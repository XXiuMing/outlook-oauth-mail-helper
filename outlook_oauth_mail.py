#!/usr/bin/env python3
import argparse
import base64
import json
import mimetypes
import re
import sys
import time
from pathlib import Path
from typing import Any, Dict
from urllib.parse import quote, urlencode

import requests

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
DEFAULT_CONFIG = Path.home() / ".outlook-oauth" / "config.json"
DEFAULT_TENANT = "common"
DEFAULT_SCOPE = "https://graph.microsoft.com/.default"
DEFAULT_AUTH_SCOPE = "offline_access https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/User.Read"
SMALL_ATTACHMENT_LIMIT = 3 * 1024 * 1024
UPLOAD_CHUNK_SIZE = 3276800
DEFAULT_SELECT = "id,subject,from,receivedDateTime,isRead,parentFolderId"

FOLDER_ALIASES = {
    "inbox": "inbox",
    "drafts": "drafts",
    "sent": "sentitems",
    "sentitems": "sentitems",
    "deleted": "deleteditems",
    "deleteditems": "deleteditems",
    "junk": "junkemail",
    "junkemail": "junkemail",
    "archive": "archive",
}


def load_config(path: Path) -> Dict[str, Any]:
    if not path.exists():
        raise SystemExit(
            f"Config not found: {path}\n"
            "Create it first. Example:\n"
            '{\n'
            '  "refresh_token": "<REFRESH_TOKEN>",\n'
            '  "client_id": "<CLIENT_ID>",\n'
            '  "tenant": "common",\n'
            '  "scope": "https://graph.microsoft.com/.default"\n'
            '}\n'
        )
    return json.loads(path.read_text(encoding="utf-8"))


def save_config(path: Path, config: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(config, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def token_endpoint(config: Dict[str, Any]) -> str:
    tenant = config.get("tenant") or DEFAULT_TENANT
    return f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"


def token_is_expired(config: Dict[str, Any], skew_seconds: int = 120) -> bool:
    expires_at = config.get("expires_at")
    if expires_at is None:
        return False
    try:
        return time.time() >= float(expires_at) - skew_seconds
    except (TypeError, ValueError):
        return False


def refresh_access_token(config_path: Path, config: Dict[str, Any]) -> Dict[str, Any]:
    refresh_token = config.get("refresh_token")
    client_id = config.get("client_id")
    if not refresh_token or not client_id:
        raise SystemExit("No usable credentials found. Need at least refresh_token + client_id in config.")

    scope = config.get("scope") or DEFAULT_SCOPE
    payload = {
        "client_id": client_id,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "scope": scope,
    }

    resp = requests.post(token_endpoint(config), data=payload, timeout=30)
    if resp.status_code >= 400:
        raise SystemExit(f"Token refresh failed: {resp.status_code} {resp.text}")

    token_data = resp.json()
    new_access_token = token_data.get("access_token")
    if not new_access_token:
        raise SystemExit(f"Token refresh response missing access_token: {token_data}")

    config["access_token"] = new_access_token
    if token_data.get("refresh_token"):
        config["refresh_token"] = token_data["refresh_token"]
    if token_data.get("scope"):
        config["scope"] = token_data["scope"]
    if token_data.get("expires_in"):
        config["expires_at"] = int(time.time()) + int(token_data["expires_in"])

    save_config(config_path, config)
    return config


def ensure_access_token(config_path: Path, config: Dict[str, Any]) -> Dict[str, Any]:
    token = config.get("access_token")
    if not token:
        return refresh_access_token(config_path, config)
    if token_is_expired(config):
        return refresh_access_token(config_path, config)
    return config


def graph_request(config_path: Path, config: Dict[str, Any], method: str, path: str, **kwargs: Any) -> requests.Response:
    config = ensure_access_token(config_path, config)
    token = config.get("access_token")
    headers = kwargs.pop("headers", {})
    headers.setdefault("Authorization", f"Bearer {token}")
    headers.setdefault("Accept", "application/json")
    resp = requests.request(method, f"{GRAPH_BASE}{path}", headers=headers, timeout=30, **kwargs)

    if resp.status_code == 401 and config.get("refresh_token") and config.get("client_id"):
        config = refresh_access_token(config_path, config)
        headers["Authorization"] = f"Bearer {config['access_token']}"
        resp = requests.request(method, f"{GRAPH_BASE}{path}", headers=headers, timeout=30, **kwargs)

    return resp


def make_recipients(addresses: list[str]) -> list[dict[str, Any]]:
    return [{"emailAddress": {"address": addr}} for addr in addresses]


def build_message_payload(subject: str, body: str, html: bool, to: list[str], cc: list[str]) -> Dict[str, Any]:
    return {
        "subject": subject,
        "body": {"contentType": "HTML" if html else "Text", "content": body},
        "toRecipients": make_recipients(to),
        "ccRecipients": make_recipients(cc),
    }


def resolve_folder(folder: str) -> str:
    return FOLDER_ALIASES.get(folder.lower(), folder)


def safe_filename(name: str, fallback: str = "message") -> str:
    name = (name or fallback).strip()
    name = re.sub(r'[\\/:*?"<>|]+', '_', name)
    return name[:180] or fallback


def upload_large_attachment(config_path: Path, config: Dict[str, Any], draft_id: str, file_path: Path) -> Dict[str, Any]:
    mime_type = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
    create_payload = {"AttachmentItem": {"attachmentType": "file", "name": file_path.name, "size": file_path.stat().st_size, "contentType": mime_type}}
    create_resp = graph_request(config_path, config, "POST", f"/me/messages/{draft_id}/attachments/createUploadSession", headers={"Content-Type": "application/json"}, json=create_payload)
    if create_resp.status_code >= 400:
        raise SystemExit(f"Create upload session failed: {create_resp.status_code} {create_resp.text}")
    upload_url = create_resp.json().get("uploadUrl")
    if not upload_url:
        raise SystemExit(f"Upload session response missing uploadUrl: {create_resp.text}")
    total_size = file_path.stat().st_size
    with file_path.open("rb") as fh:
        start = 0
        while start < total_size:
            fh.seek(start)
            chunk = fh.read(min(UPLOAD_CHUNK_SIZE, total_size - start))
            if not chunk:
                break
            end = start + len(chunk) - 1
            headers = {"Content-Length": str(len(chunk)), "Content-Range": f"bytes {start}-{end}/{total_size}"}
            resp = requests.put(upload_url, headers=headers, data=chunk, timeout=60)
            if resp.status_code not in (200, 201, 202):
                raise SystemExit(f"Chunk upload failed: {resp.status_code} {resp.text}")
            payload = resp.json() if resp.text.strip() else {}
            next_ranges = payload.get("nextExpectedRanges") if isinstance(payload, dict) else None
            if next_ranges:
                start = int(next_ranges[0].split("-")[0])
                continue
            return payload or {"status": "uploaded"}
    return {"status": "uploaded"}


def list_attachments(config_path: Path, config: Dict[str, Any], message_id: str) -> Dict[str, Any]:
    resp = graph_request(config_path, config, "GET", f"/me/messages/{message_id}/attachments", params={"$select": "id,name,contentType,size,isInline"})
    if resp.status_code >= 400:
        raise SystemExit(resp.text)
    return resp.json()


def download_attachment_to(config_path: Path, config: Dict[str, Any], message_id: str, attachment_id: str, out_dir: Path) -> Path:
    file_resp = graph_request(config_path, config, "GET", f"/me/messages/{message_id}/attachments/{attachment_id}")
    if file_resp.status_code >= 400:
        raise SystemExit(file_resp.text)
    data = file_resp.json()
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / safe_filename(data.get("name") or attachment_id, fallback=attachment_id)
    if data.get("contentBytes"):
        out_path.write_bytes(base64.b64decode(data["contentBytes"]))
        return out_path
    value_resp = graph_request(config_path, config, "GET", f"/me/messages/{message_id}/attachments/{attachment_id}/$value")
    if value_resp.status_code >= 400:
        raise SystemExit(value_resp.text)
    out_path.write_bytes(value_resp.content)
    return out_path


def cmd_inbox(config_path: Path, config: Dict[str, Any], args: argparse.Namespace) -> int:
    params = {"$top": args.limit, "$select": DEFAULT_SELECT, "$orderby": "receivedDateTime desc"}
    if args.unread:
        params["$filter"] = "isRead eq false"
    resp = graph_request(config_path, config, "GET", "/me/messages", params=params)
    print(json.dumps(resp.json(), ensure_ascii=False, indent=2) if resp.status_code < 400 else resp.text, file=sys.stdout if resp.status_code < 400 else sys.stderr)
    return 0 if resp.status_code < 400 else 1


def cmd_read(config_path: Path, config: Dict[str, Any], args: argparse.Namespace) -> int:
    resp = graph_request(config_path, config, "GET", f"/me/messages/{args.message_id}", params={"$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,isRead,parentFolderId"})
    print(json.dumps(resp.json(), ensure_ascii=False, indent=2) if resp.status_code < 400 else resp.text, file=sys.stdout if resp.status_code < 400 else sys.stderr)
    return 0 if resp.status_code < 400 else 1


def cmd_search(config_path: Path, config: Dict[str, Any], args: argparse.Namespace) -> int:
    path = f"/me/mailFolders/{quote(resolve_folder(args.folder), safe='')}/messages"
    resp = graph_request(config_path, config, "GET", path, headers={"ConsistencyLevel": "eventual"}, params={"$top": args.limit, "$select": DEFAULT_SELECT, "$search": f'"{args.query}"'})
    print(json.dumps(resp.json(), ensure_ascii=False, indent=2) if resp.status_code < 400 else resp.text, file=sys.stdout if resp.status_code < 400 else sys.stderr)
    return 0 if resp.status_code < 400 else 1


def cmd_list_drafts(config_path: Path, config: Dict[str, Any], args: argparse.Namespace) -> int:
    resp = graph_request(config_path, config, "GET", "/me/mailFolders/drafts/messages", params={"$top": args.limit, "$select": DEFAULT_SELECT, "$orderby": "receivedDateTime desc"})
    print(json.dumps(resp.json(), ensure_ascii=False, indent=2) if resp.status_code < 400 else resp.text, file=sys.stdout if resp.status_code < 400 else sys.stderr)
    return 0 if resp.status_code < 400 else 1


def cmd_send(config_path: Path, config: Dict[str, Any], args: argparse.Namespace) -> int:
    resp = graph_request(config_path, config, "POST", "/me/sendMail", headers={"Content-Type": "application/json"}, json={"message": build_message_payload(args.subject, args.body, args.html, args.to, args.cc), "saveToSentItems": True})
    print("Mail sent." if resp.status_code in (200, 202) else resp.text, file=sys.stdout if resp.status_code in (200, 202) else sys.stderr)
    return 0 if resp.status_code in (200, 202) else 1


def cmd_draft(config_path: Path, config: Dict[str, Any], args: argparse.Namespace) -> int:
    resp = graph_request(config_path, config, "POST", "/me/messages", headers={"Content-Type": "application/json"}, json=build_message_payload(args.subject, args.body, args.html, args.to, args.cc))
    if resp.status_code >= 400:
        print(resp.text, file=sys.stderr); return 1
    data = resp.json(); print(json.dumps({"id": data.get("id"), "subject": data.get("subject")}, ensure_ascii=False, indent=2)); return 0


def cmd_reply(config_path: Path, config: Dict[str, Any], args: argparse.Namespace) -> int:
    endpoint = f"/me/messages/{args.message_id}/createReplyAll" if args.all else f"/me/messages/{args.message_id}/createReply"
    resp = graph_request(config_path, config, "POST", endpoint)
    if resp.status_code >= 400:
        print(resp.text, file=sys.stderr); return 1
    draft_id = resp.json().get("id")
    if not draft_id:
        print(resp.text, file=sys.stderr); return 1
    patch_payload = {"body": {"contentType": "HTML" if args.html else "Text", "content": args.body}}
    if args.to: patch_payload["toRecipients"] = make_recipients(args.to)
    if args.cc: patch_payload["ccRecipients"] = make_recipients(args.cc)
    patch_resp = graph_request(config_path, config, "PATCH", f"/me/messages/{draft_id}", headers={"Content-Type": "application/json"}, json=patch_payload)
    if patch_resp.status_code >= 400:
        print(patch_resp.text, file=sys.stderr); return 1
    if args.send_now:
        send_resp = graph_request(config_path, config, "POST", f"/me/messages/{draft_id}/send")
        print("Reply sent." if send_resp.status_code in (200, 202) else send_resp.text, file=sys.stdout if send_resp.status_code in (200, 202) else sys.stderr)
        return 0 if send_resp.status_code in (200, 202) else 1
    print(json.dumps({"draft_id": draft_id, "status": "reply draft created"}, ensure_ascii=False, indent=2)); return 0


def cmd_attach(config_path: Path, config: Dict[str, Any], args: argparse.Namespace) -> int:
    file_path = Path(args.file).expanduser()
    if not file_path.exists() or not file_path.is_file(): raise SystemExit(f"Attachment file not found: {file_path}")
    size = file_path.stat().st_size
    mime_type = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
    if size <= SMALL_ATTACHMENT_LIMIT:
        payload = {"@odata.type": "#microsoft.graph.fileAttachment", "name": file_path.name, "contentType": mime_type, "contentBytes": base64.b64encode(file_path.read_bytes()).decode("ascii")}
        resp = graph_request(config_path, config, "POST", f"/me/messages/{args.draft_id}/attachments", headers={"Content-Type": "application/json"}, json=payload)
        print(json.dumps(resp.json(), ensure_ascii=False, indent=2) if resp.status_code < 400 else resp.text, file=sys.stdout if resp.status_code < 400 else sys.stderr)
        return 0 if resp.status_code < 400 else 1
    print(json.dumps(upload_large_attachment(config_path, config, args.draft_id, file_path), ensure_ascii=False, indent=2)); return 0


def cmd_send_draft(config_path: Path, config: Dict[str, Any], args: argparse.Namespace) -> int:
    resp = graph_request(config_path, config, "POST", f"/me/messages/{args.draft_id}/send")
    print("Draft sent." if resp.status_code in (200, 202) else resp.text, file=sys.stdout if resp.status_code in (200, 202) else sys.stderr)
    return 0 if resp.status_code in (200, 202) else 1


def cmd_delete(config_path: Path, config: Dict[str, Any], args: argparse.Namespace) -> int:
    resp = graph_request(config_path, config, "DELETE", f"/me/messages/{args.message_id}")
    print("Message deleted." if resp.status_code in (200, 202, 204) else resp.text, file=sys.stdout if resp.status_code in (200, 202, 204) else sys.stderr)
    return 0 if resp.status_code in (200, 202, 204) else 1


def cmd_folders(config_path: Path, config: Dict[str, Any], args: argparse.Namespace) -> int:
    resp = graph_request(config_path, config, "GET", "/me/mailFolders", params={"$top": args.limit, "$select": "id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount"})
    print(json.dumps(resp.json(), ensure_ascii=False, indent=2) if resp.status_code < 400 else resp.text, file=sys.stdout if resp.status_code < 400 else sys.stderr)
    return 0 if resp.status_code < 400 else 1


def cmd_mark(config_path: Path, config: Dict[str, Any], args: argparse.Namespace) -> int:
    resp = graph_request(config_path, config, "PATCH", f"/me/messages/{args.message_id}", headers={"Content-Type": "application/json"}, json={"isRead": args.read})
    print("Message updated." if resp.status_code < 400 else resp.text, file=sys.stdout if resp.status_code < 400 else sys.stderr)
    return 0 if resp.status_code < 400 else 1


def cmd_list_attachments(config_path: Path, config: Dict[str, Any], args: argparse.Namespace) -> int:
    print(json.dumps(list_attachments(config_path, config, args.message_id), ensure_ascii=False, indent=2)); return 0


def cmd_download_attachment(config_path: Path, config: Dict[str, Any], args: argparse.Namespace) -> int:
    print(str(download_attachment_to(config_path, config, args.message_id, args.attachment_id, Path(args.outdir).expanduser()))); return 0


def cmd_download_all_attachments(config_path: Path, config: Dict[str, Any], args: argparse.Namespace) -> int:
    data = list_attachments(config_path, config, args.message_id)
    out_dir = Path(args.outdir).expanduser(); saved = []
    for item in data.get("value", []):
        item_type = item.get("@odata.type", "")
        if item_type and item_type != "#microsoft.graph.fileAttachment":
            continue
        saved.append(str(download_attachment_to(config_path, config, args.message_id, item["id"], out_dir)))
    print(json.dumps({"saved": saved}, ensure_ascii=False, indent=2)); return 0


def cmd_save_body(config_path: Path, config: Dict[str, Any], args: argparse.Namespace) -> int:
    resp = graph_request(config_path, config, "GET", f"/me/messages/{args.message_id}", params={"$select": "subject,body,bodyPreview,receivedDateTime,from"})
    if resp.status_code >= 400:
        print(resp.text, file=sys.stderr); return 1
    data = resp.json(); out_dir = Path(args.outdir).expanduser(); out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{safe_filename(data.get('subject') or args.message_id)}{'.html' if args.html else '.txt'}"
    body = (data.get("body") or {}).get("content", "")
    if not args.html: body = re.sub(r"<[^>]+>", "", body)
    out_path.write_text(body, encoding="utf-8"); print(str(out_path)); return 0


def cmd_move(config_path: Path, config: Dict[str, Any], args: argparse.Namespace) -> int:
    resp = graph_request(config_path, config, "POST", f"/me/messages/{args.message_id}/move", headers={"Content-Type": "application/json"}, json={"destinationId": resolve_folder(args.destination)})
    print(json.dumps(resp.json(), ensure_ascii=False, indent=2) if resp.status_code < 400 else resp.text, file=sys.stdout if resp.status_code < 400 else sys.stderr)
    return 0 if resp.status_code < 400 else 1


def cmd_auth_url(config: Dict[str, Any], args: argparse.Namespace) -> int:
    client_id = config.get("client_id")
    if not client_id: raise SystemExit("Missing client_id in config")
    tenant = config.get("tenant") or DEFAULT_TENANT
    scope = config.get("auth_scope") or DEFAULT_AUTH_SCOPE
    redirect_uri = args.redirect_uri or config.get("redirect_uri") or "http://localhost:8080"
    print(f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize?" + urlencode({"client_id": client_id, "response_type": "code", "redirect_uri": redirect_uri, "response_mode": "query", "scope": scope}))
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Outlook OAuth / Microsoft Graph mail helper")
    parser.add_argument("--config", default=str(DEFAULT_CONFIG), help="Path to JSON config file")
    sub = parser.add_subparsers(dest="command", required=True)
    inbox = sub.add_parser("inbox", help="List inbox messages"); inbox.add_argument("-n", "--limit", type=int, default=10); inbox.add_argument("-u", "--unread", action="store_true")
    read = sub.add_parser("read", help="Read a single message"); read.add_argument("message_id")
    search = sub.add_parser("search", help="Search messages in a folder"); search.add_argument("query"); search.add_argument("--folder", default="inbox"); search.add_argument("-n", "--limit", type=int, default=10)
    drafts = sub.add_parser("drafts", help="List draft messages"); drafts.add_argument("-n", "--limit", type=int, default=10)
    send = sub.add_parser("send", help="Send a message"); send.add_argument("--to", nargs="+", required=True); send.add_argument("--cc", nargs="*", default=[]); send.add_argument("--subject", required=True); send.add_argument("--body", required=True); send.add_argument("--html", action="store_true")
    draft = sub.add_parser("draft", help="Create a draft message"); draft.add_argument("--to", nargs="+", required=True); draft.add_argument("--cc", nargs="*", default=[]); draft.add_argument("--subject", required=True); draft.add_argument("--body", required=True); draft.add_argument("--html", action="store_true")
    reply = sub.add_parser("reply", help="Create a reply draft or send reply"); reply.add_argument("message_id"); reply.add_argument("--body", required=True); reply.add_argument("--to", nargs="*", default=[]); reply.add_argument("--cc", nargs="*", default=[]); reply.add_argument("--html", action="store_true"); reply.add_argument("--all", action="store_true"); reply.add_argument("--send-now", action="store_true")
    attach = sub.add_parser("attach", help="Attach a file to an existing draft"); attach.add_argument("draft_id"); attach.add_argument("file")
    send_draft = sub.add_parser("send-draft", help="Send an existing draft"); send_draft.add_argument("draft_id")
    delete = sub.add_parser("delete", help="Delete a message"); delete.add_argument("message_id")
    folders = sub.add_parser("folders", help="List mail folders"); folders.add_argument("-n", "--limit", type=int, default=50)
    mark = sub.add_parser("mark", help="Mark a message read or unread"); mark.add_argument("message_id"); grp = mark.add_mutually_exclusive_group(required=True); grp.add_argument("--read", action="store_true"); grp.add_argument("--unread", action="store_true")
    attachments = sub.add_parser("attachments", help="List attachments on a message"); attachments.add_argument("message_id")
    dl = sub.add_parser("download-attachment", help="Download one attachment"); dl.add_argument("message_id"); dl.add_argument("attachment_id"); dl.add_argument("--outdir", default=".")
    dla = sub.add_parser("download-all-attachments", help="Download all file attachments"); dla.add_argument("message_id"); dla.add_argument("--outdir", default=".")
    save = sub.add_parser("save-body", help="Save message body to a file"); save.add_argument("message_id"); save.add_argument("--outdir", default="."); save.add_argument("--html", action="store_true")
    move = sub.add_parser("move", help="Move a message to another folder"); move.add_argument("message_id"); move.add_argument("destination")
    auth = sub.add_parser("auth-url", help="Print OAuth authorize URL"); auth.add_argument("--redirect-uri", default=None)
    return parser


def main() -> int:
    parser = build_parser(); args = parser.parse_args(); config_path = Path(args.config).expanduser(); config = load_config(config_path)
    cmds = {
        "inbox": cmd_inbox, "read": cmd_read, "search": cmd_search, "drafts": cmd_list_drafts, "send": cmd_send,
        "draft": cmd_draft, "reply": cmd_reply, "attach": cmd_attach, "send-draft": cmd_send_draft,
        "delete": cmd_delete, "folders": cmd_folders, "mark": cmd_mark, "attachments": cmd_list_attachments,
        "download-attachment": cmd_download_attachment, "download-all-attachments": cmd_download_all_attachments,
        "save-body": cmd_save_body, "move": cmd_move, "auth-url": lambda p, c, a: cmd_auth_url(c, a),
    }
    if args.command == "mark": args.read = True if args.read else False
    return cmds[args.command](config_path, config, args)


if __name__ == "__main__":
    raise SystemExit(main())
