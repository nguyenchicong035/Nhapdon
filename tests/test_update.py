from hashlib import sha256
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
from unittest.mock import patch
import updater


class UpdateTests(TestCase):
    def test_version_and_verified_download(self):
        tag = "v99.0.0"
        data = b"test asset"
        url = f"https://github.com/{updater.REPO}/releases/download/{tag}/Nhapdon.exe"
        asset = dict(browser_download_url=url, size=len(data), digest="sha256:"+sha256(data).hexdigest())
        release = updater.newer_release(dict(tag_name=tag, assets=[asset]))
        self.assertIsNone(updater.newer_release(dict(tag_name="v"+updater.VERSION)))
        with TemporaryDirectory() as folder:
            with patch("updater.urlopen", return_value=BytesIO(data)):
                path = updater.download(release, folder)
                self.assertEqual(path.read_bytes(), data)
                path.unlink()
            with patch("updater.urlopen", return_value=BytesIO(b"x"*len(data))):
                with self.assertRaises(ValueError):
                    updater.download(release, folder)
            self.assertEqual(list(Path(folder).iterdir()), [])
