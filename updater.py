"""Download a stable Nhapdon release and replace the Windows executable on exit."""
import base64
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
from tempfile import NamedTemporaryFile
from urllib.request import Request, urlopen

VERSION = "1.0.6"
REPO = "nguyenchicong035/Nhapdon"


def newer_release(data):
    tag = data["tag_name"]
    if data.get("draft") or data.get("prerelease"):
        return None
    if not re.fullmatch(r"v\d+\.\d+\.\d+", tag):
        raise ValueError("Không đọc được số phiên bản phát hành.")
    if tuple(map(int, tag[1:].split("."))) <= tuple(map(int, VERSION.split("."))):
        return None
    url = f"https://github.com/{REPO}/releases/download/{tag}/Nhapdon.exe"
    asset = next((a for a in data["assets"] if a["browser_download_url"] == url), None)
    if not asset or not re.fullmatch(r"sha256:[0-9a-f]{64}", asset.get("digest") or "") or asset["size"] <= 0:
        raise ValueError("Bản phát hành chưa có Nhapdon.exe kèm mã kiểm tra SHA-256.")
    return tag, url, asset["size"], asset["digest"][7:]


def latest():
    req = Request(f"https://api.github.com/repos/{REPO}/releases/latest",
                  headers={"User-Agent": f"Nhapdon/{VERSION}", "Accept": "application/vnd.github+json"})
    with urlopen(req, timeout=20) as response:
        return newer_release(json.load(response))


def download(release, directory):
    _, url, expected, checksum = release
    digest = hashlib.sha256()
    size = 0
    with NamedTemporaryFile(dir=directory, prefix=".Nhapdon-update-", suffix=".exe", delete=False) as out:
        path = Path(out.name)
    try:
        with urlopen(url, timeout=60) as response, path.open("wb") as out:
            while block := response.read(1024 * 1024):
                size += len(block)
                if size > expected:
                    raise ValueError("Kích thước bản tải về không khớp.")
                digest.update(block)
                out.write(block)
        if size != expected or digest.hexdigest() != checksum:
            raise ValueError("Bản tải về không khớp với SHA-256 hoặc kích thước trên GitHub.")
        return path
    except Exception:
        path.unlink(missing_ok=True)
        raise


# Pass paths as environment data so spaces and PowerShell metacharacters are literal.
SWAP = r"""
$ErrorActionPreference = 'Stop'
$new = $env:NH_UPDATE_FILE
$target = $env:NH_UPDATE_TARGET
$backup = $new + '.previous'
try {
    $test = Start-Process -FilePath $new -ArgumentList '--self-test' -PassThru
    if (-not $test.WaitForExit(120000) -or $test.ExitCode -ne 0) { throw 'Bản mới không khởi động được.' }
    for ($i = 0; $i -lt 60; $i++) {
        try { [IO.File]::Replace($new, $target, $backup); break }
        catch [IO.IOException] { if ($i -eq 59) { throw }; Start-Sleep -Seconds 1 }
    }
    try { Start-Process -FilePath $target }
    catch { [IO.File]::Replace($backup, $target, $null); throw }
    Remove-Item -LiteralPath $backup -ErrorAction SilentlyContinue
} catch {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Cập nhật Nhapdon')
} finally {
    Remove-Item -LiteralPath $new -ErrorAction SilentlyContinue
}
"""


def install_on_exit(path, target):
    env = dict(os.environ, NH_UPDATE_FILE=str(path), NH_UPDATE_TARGET=str(target),
               PYINSTALLER_RESET_ENVIRONMENT="1")
    encoded = base64.b64encode(SWAP.encode("utf-16le")).decode("ascii")
    shell = Path(os.environ["SystemRoot"]) / "System32/WindowsPowerShell/v1.0/powershell.exe"
    return subprocess.Popen([str(shell), "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded],
                            env=env, creationflags=subprocess.CREATE_NO_WINDOW)
