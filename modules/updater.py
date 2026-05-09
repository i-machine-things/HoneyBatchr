"""Background update checker, downloader, and update dialog for HoneyBatchr."""

import html
import json
import os
import platform
import subprocess
import tempfile
import urllib.error
import urllib.request
import webbrowser
from pathlib import Path

from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtWidgets import (
    QApplication, QCheckBox, QDialog, QDialogButtonBox,
    QLabel, QMessageBox, QProgressDialog, QVBoxLayout,
)

from modules.config import update_config_value, load_config

_GITHUB_REPO = "i-machine-things/HoneyBatchr"


def get_app_version() -> str:
    try:
        from modules._version import __version__ as _v
        if _v and _v != "dev":
            return _v
    except ImportError:
        pass
    try:
        result = subprocess.run(
            ["git", "describe", "--tags", "--abbrev=0"],
            capture_output=True, text=True, timeout=3,
            cwd=Path(__file__).parent.parent,
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip()
    except Exception:
        pass
    return "dev"


APP_VERSION = get_app_version()


def _version_tuple(v: str) -> tuple:
    try:
        numeric = v.lstrip("v").split("-")[0]
        return tuple(int(x) for x in numeric.split(".")[:3])
    except ValueError:
        return (0, 0, 0)


class UpdateChecker(QThread):
    """Queries the GitHub releases API on a background thread."""
    update_available = pyqtSignal(str, str, str)  # (tag, html_url, asset_url)
    up_to_date = pyqtSignal()
    check_failed = pyqtSignal(str)

    def run(self):
        try:
            url = f"https://api.github.com/repos/{_GITHUB_REPO}/releases/latest"
            req = urllib.request.Request(
                url,
                headers={"Accept": "application/vnd.github+json", "User-Agent": "HoneyBatchr"},
            )
            with urllib.request.urlopen(req, timeout=10) as resp:  # nosec B310
                data = json.loads(resp.read().decode())
            tag = data.get("tag_name", "")
            html_url = data.get("html_url", "")
            if tag and _version_tuple(tag) > _version_tuple(APP_VERSION):
                asset_url = ""
                for asset in data.get("assets", []):
                    if asset.get("name", "").lower().endswith(".exe"):
                        asset_url = asset.get("browser_download_url", "")
                        break
                self.update_available.emit(tag, html_url, asset_url)
            else:
                self.up_to_date.emit()
        except (urllib.error.URLError, OSError, json.JSONDecodeError, ValueError) as exc:
            self.check_failed.emit(str(exc))


class UpdateDownloader(QThread):
    """Downloads a release asset to a temp file on a background thread."""
    progress = pyqtSignal(int, int)  # (bytes_done, total_bytes)
    download_finished = pyqtSignal(str)  # dest_path on success
    cancelled = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, asset_url: str, dest_path: str, parent=None):
        super().__init__(parent)
        self._asset_url = asset_url
        self._dest_path = dest_path

    def run(self):
        try:
            if not self._asset_url.startswith("https://"):
                self.error.emit("Invalid asset URL scheme")
                return
            req = urllib.request.Request(
                self._asset_url,
                headers={"User-Agent": "HoneyBatchr"},
            )
            cancelled = False
            with urllib.request.urlopen(req, timeout=60) as resp:  # nosec B310
                total = int(resp.headers.get("Content-Length", 0))
                done = 0
                with open(self._dest_path, "wb") as f:
                    while True:
                        if self.isInterruptionRequested():
                            cancelled = True
                            break
                        buf = resp.read(65536)
                        if not buf:
                            break
                        f.write(buf)
                        done += len(buf)
                        self.progress.emit(done, total)
            if cancelled:
                try:
                    os.remove(self._dest_path)
                except OSError:
                    pass
                self.cancelled.emit()
                return
            if total > 0 and done != total:
                try:
                    os.remove(self._dest_path)
                except OSError:
                    pass
                self.error.emit(f"Download truncated: received {done} of {total} bytes")
                return
            self.download_finished.emit(self._dest_path)
        except Exception as exc:
            try:
                os.remove(self._dest_path)
            except OSError:
                pass
            self.error.emit(str(exc))


class UpdateDialog(QDialog):
    """Non-blocking update-available prompt."""

    def __init__(self, latest_version: str, release_url: str, asset_url: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Update Available")
        self.setMinimumWidth(340)
        self._latest_version = latest_version
        self._release_url = release_url
        self._asset_url = asset_url

        self._is_flatpak = platform.system() == 'Linux' and bool(os.getenv('FLATPAK_ID'))
        self._can_download = platform.system() == 'Windows' and bool(asset_url)

        layout = QVBoxLayout(self)

        label = QLabel(f"<b>{html.escape(latest_version)}</b> is available. Upgrade now?")
        label.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(label)

        self._disable_cb = QCheckBox("Don't show update notifications")
        layout.addWidget(self._disable_cb)

        buttons = QDialogButtonBox()
        ok_label = "Download && Install" if self._can_download else "Update Now"
        buttons.addButton(ok_label, QDialogButtonBox.ButtonRole.AcceptRole)
        buttons.addButton("Skip This Version", QDialogButtonBox.ButtonRole.RejectRole)
        buttons.accepted.connect(self._on_ok)
        buttons.rejected.connect(self._on_later)
        layout.addWidget(buttons)

    def _save_disabled(self) -> None:
        if self._disable_cb.isChecked():
            update_config_value('updates_notifications_disabled', True)

    def _on_ok(self) -> None:
        self._save_disabled()
        if self._can_download:
            self._start_download()
        else:
            webbrowser.open(self._release_url)
            self.accept()

    def _start_download(self) -> None:
        safe_ver = ''.join(c if c.isalnum() or c in '._-' else '_' for c in self._latest_version)
        fd, dest = tempfile.mkstemp(suffix=".exe", prefix=f"HoneyBatchr-{safe_ver}-")
        os.close(fd)

        progress_dlg = QProgressDialog(
            f"Downloading {self._latest_version}...", "Cancel", 0, 100, self,
        )
        progress_dlg.setWindowTitle("Downloading Update")
        progress_dlg.setWindowModality(Qt.WindowModality.WindowModal)
        progress_dlg.setMinimumDuration(0)
        progress_dlg.setValue(0)

        downloader = UpdateDownloader(self._asset_url, dest, self)
        self._downloader = downloader
        progress_dlg.canceled.connect(downloader.requestInterruption)

        def _on_progress(done: int, total: int) -> None:
            progress_dlg.setValue(int(done * 100 / total) if total else 50)

        def _on_done(path: str) -> None:
            progress_dlg.close()
            try:
                with open(path, 'rb') as _f:
                    if _f.read(2) != b'MZ':
                        raise ValueError("Not a valid Windows executable (bad MZ header)")
                    _f.seek(0x3C)
                    e_lfanew = int.from_bytes(_f.read(4), 'little')
                    _f.seek(e_lfanew)
                    if _f.read(4) != b'PE\x00\x00':
                        raise ValueError("Not a valid Windows executable (bad PE signature)")
            except (OSError, ValueError) as exc:
                QMessageBox.critical(
                    self, "Verification Failed",
                    f"The downloaded installer failed verification:\n{exc}\n\n"
                    "Opening the releases page instead.",
                )
                try:
                    os.remove(path)
                except OSError:
                    pass
                webbrowser.open(self._release_url)
                self.accept()
                return
            reply = QMessageBox.question(
                self, "Ready to Install",
                f"{self._latest_version} downloaded.\n\n"
                "Honey Batchr will close and the installer will launch.\nReady to install?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.Yes:
                try:
                    subprocess.Popen([path])
                except OSError as exc:
                    QMessageBox.critical(
                        self, "Launch Failed",
                        f"Could not launch the installer:\n{exc}\n\nInstaller saved at:\n{path}",
                    )
                    self.accept()
                    return
                QApplication.quit()
            self.accept()

        def _on_cancelled() -> None:
            progress_dlg.close()

        def _on_error(msg: str) -> None:
            progress_dlg.close()
            QMessageBox.warning(
                self, "Download Failed",
                f"Could not download the update:\n{msg}\n\nOpening the releases page instead.",
            )
            webbrowser.open(self._release_url)
            self.accept()

        downloader.progress.connect(_on_progress)
        downloader.download_finished.connect(_on_done)
        downloader.cancelled.connect(_on_cancelled)
        downloader.error.connect(_on_error)
        downloader.download_finished.connect(downloader.deleteLater)
        downloader.cancelled.connect(downloader.deleteLater)
        downloader.error.connect(downloader.deleteLater)
        downloader.start()

    def _on_later(self) -> None:
        self._save_disabled()
        update_config_value('updates_skipped_version', self._latest_version)
        self.reject()


def flatpak_dns_fix() -> None:
    """Patch socket.getaddrinfo for Flatpak sandboxes where DNS fails.

    On systemd-resolved distros, /etc/resolv.conf contains 'nameserver 127.0.0.53'
    (the stub resolver). That address only listens on the host loopback and is
    unreachable inside the Flatpak network namespace, causing Errno -3 on every
    name lookup. This reads the real upstream nameservers and falls back to a
    direct UDP DNS query when the system resolver fails.
    """
    if not os.getenv('FLATPAK_ID'):
        return
    import re
    import secrets
    import socket
    import struct

    try:
        socket.getaddrinfo('github.com', 443, socket.AF_INET)
        return
    except (socket.gaierror, OSError):
        pass

    nameservers: list = []
    for path in ('/run/systemd/resolve/resolv.conf',
                 '/run/host/etc/resolv.conf',
                 '/etc/resolv.conf'):
        try:
            with open(path) as _f:
                for _line in _f:
                    _m = re.match(r'^nameserver\s+(\S+)', _line)
                    if _m:
                        _ns = _m.group(1)
                        if not _ns.startswith('127.') and _ns != '::1':
                            nameservers.append(_ns)
        except OSError:
            continue
        if nameservers:
            break

    if not nameservers:
        nameservers = ['1.1.1.1', '8.8.8.8']

    def _dns_a(host: str, ns: str) -> list:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.settimeout(3)
        try:
            qid = secrets.randbelow(65536)
            labels = b''.join(bytes([len(p)]) + p.encode()
                              for p in host.rstrip('.').split('.')) + b'\x00'
            pkt = (struct.pack('!HHHHHH', qid, 0x0100, 1, 0, 0, 0)
                   + labels + struct.pack('!HH', 1, 1))
            sock.sendto(pkt, (ns, 53))
            resp = sock.recv(512)
        finally:
            sock.close()
        if len(resp) < 12 or struct.unpack('!H', resp[:2])[0] != qid:
            return []
        ancount = struct.unpack('!H', resp[6:8])[0]
        pos = 12
        while pos < len(resp) and resp[pos]:
            pos += resp[pos] + 1
        if pos + 5 > len(resp):
            return []
        pos += 5
        ips: list = []
        for _ in range(ancount):
            if resp[pos] & 0xC0 == 0xC0:
                pos += 2
            else:
                while resp[pos]:
                    pos += resp[pos] + 1
                pos += 1
            if pos + 10 > len(resp):
                break
            rtype, _, _, rdlen = struct.unpack('!HHIH', resp[pos:pos + 10])
            pos += 10
            if rtype == 1 and rdlen == 4:
                ips.append(socket.inet_ntoa(resp[pos:pos + 4]))
            pos += rdlen
        return ips

    _orig_getaddrinfo = socket.getaddrinfo
    _ns_list = nameservers[:]

    def _getaddrinfo(host, port, family=0, socktype=0, proto=0, flags=0):
        try:
            return _orig_getaddrinfo(host, port, family, socktype, proto, flags)
        except (socket.gaierror, OSError):
            pass
        p = port if isinstance(port, int) else 0
        for _ns in _ns_list:
            try:
                for _ip in _dns_a(host, _ns):
                    return [(socket.AF_INET, socket.SOCK_STREAM, 6, '', (_ip, p)),
                            (socket.AF_INET, socket.SOCK_DGRAM, 17, '', (_ip, p))]
            except Exception:
                continue
        return _orig_getaddrinfo(host, port, family, socktype, proto, flags)

    socket.getaddrinfo = _getaddrinfo
