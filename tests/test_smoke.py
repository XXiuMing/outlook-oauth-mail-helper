from pathlib import Path
import subprocess
import sys
import unittest


class SmokeTests(unittest.TestCase):
    def test_help_runs(self):
        repo_root = Path(__file__).resolve().parents[1]
        script = repo_root / "outlook_oauth_mail.py"
        result = subprocess.run(
            [sys.executable, str(script), "--help"],
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertEqual(result.returncode, 0)
        self.assertTrue("outlook" in result.stdout.lower() or "mail" in result.stdout.lower())


if __name__ == "__main__":
    unittest.main()
