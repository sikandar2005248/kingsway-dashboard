"""
Kingsway dashboard publisher - runs on GitHub (no laptop needed).

GitHub runs this every 15 minutes (see .github/workflows/update-data.yml).
It downloads the Kingsway workbook from its OneDrive share link, works out
the dashboard figures (same logic as the laptop version, kw_data.py),
encrypts them with the dashboard password (AES-256-GCM, PBKDF2-SHA256) and
writes data.enc. The workflow then commits data.enc if anything changed.

Settings (repo Settings > Secrets and variables > Actions):
  secret   ONEDRIVE_URL        the OneDrive "Anyone with the link" share link
  secret   DASHBOARD_PASSWORD  the password phones use
  variable SEMESTER_STARTS     optional, e.g.  B1=2024-03, B2=2024-09
"""
import base64
import datetime as dt
import email.utils
import hashlib
import hmac
import json
import os
import subprocess
import sys
import tempfile
import time
import urllib.parse
import urllib.request
from zoneinfo import ZoneInfo

from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import kw_data  # noqa: E402

ITERATIONS = 250000
FORCE_EVERY_H = 6          # republish at least this often so phones know updates are running
TZ = ZoneInfo("Asia/Karachi")
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130 Safari/537.36"


def out(msg):
    print(msg, flush=True)


def set_output(name, value):
    path = os.environ.get("GITHUB_OUTPUT")
    if path:
        with open(path, "a", encoding="utf-8") as f:
            f.write("%s=%s\n" % (name, value))


# ------------------------------------------------------------ download
def _with_param(url, key, value):
    p = urllib.parse.urlsplit(url)
    q = urllib.parse.parse_qsl(p.query, keep_blank_values=True)
    q = [(k, v) for k, v in q if k != key] + [(key, value)]
    return urllib.parse.urlunsplit((p.scheme, p.netloc, p.path, urllib.parse.urlencode(q), p.fragment))


def _share_token(url):
    b = base64.b64encode(url.encode("utf-8")).decode("ascii")
    return "u!" + b.rstrip("=").replace("/", "_").replace("+", "-")


def _get(url, headers=None, data=None):
    h = {"User-Agent": UA, "Accept": "*/*"}
    h.update(headers or {})
    req = urllib.request.Request(url, headers=h, data=data, method="POST" if data is not None else "GET")
    with urllib.request.urlopen(req, timeout=90) as r:
        return r.read(), r.headers.get("Last-Modified"), r.geturl()


BADGER_APP_ID = "5cbed6ac-a083-4e14-b191-b4ba07653de2"   # the OneDrive website's own app id for shared links


def _download_badger(share_url):
    """The way onedrive.live.com itself opens an 'Anyone with the link' file: it gets an
    anonymous guest token ("Badger") and uses it to read the shared item."""
    body, _, _ = _get("https://api-badgerp.svc.ms/v1.0/token",
                      {"Content-Type": "application/json"},
                      json.dumps({"appId": BADGER_APP_ID}).encode("utf-8"))
    token = json.loads(body.decode("utf-8"))["token"]
    auth = {"Authorization": "Badger " + token, "Prefer": "autoredeem"}
    share = _share_token(share_url)
    modified = None
    last_error = None
    for host in ("https://my.microsoftpersonalcontent.com/_api/v2.0", "https://api.onedrive.com/v1.0"):
        try:
            meta_raw, _, _ = _get("%s/shares/%s/driveitem" % (host, share), dict(auth, Accept="application/json"))
            meta = json.loads(meta_raw.decode("utf-8"))
            if meta.get("lastModifiedDateTime"):
                when = dt.datetime.fromisoformat(meta["lastModifiedDateTime"].replace("Z", "+00:00"))
                modified = email.utils.format_datetime(when.astimezone(dt.timezone.utc), usegmt=True)
            link = meta.get("@content.downloadUrl") or meta.get("@microsoft.graph.downloadUrl")
            if link:
                data, _, _ = _get(link)
            else:
                data, _, _ = _get("%s/shares/%s/driveitem/content" % (host, share), auth)
            if data[:2] == b"PK":
                return data, modified
            last_error = "got something that isn't an Excel file"
        except Exception as e:  # noqa: BLE001
            last_error = e
    raise RuntimeError(last_error)


def download(share_url):
    """Try the known ways of downloading a OneDrive share link without signing in."""
    share_url = share_url.strip()
    errors = []
    try:
        body, modified = _download_badger(share_url)
        out("Downloaded workbook via OneDrive guest access (%d bytes)" % len(body))
        return body, modified
    except Exception as e:  # noqa: BLE001
        errors.append("OneDrive guest access: %s" % e)
    tries = [
        ("share link + download=1", _with_param(share_url, "download", "1")),
        ("OneDrive shares API", "https://api.onedrive.com/v1.0/shares/%s/root/content" % _share_token(share_url)),
    ]
    for name, url in tries:
        try:
            body, modified, final = _get(url)
            if body[:2] == b"PK":
                out("Downloaded workbook via %s (%d bytes)" % (name, len(body)))
                return body, modified
            # Some links land on a viewer page; ask that page for a download as well.
            if final and final != url:
                body, modified, _ = _get(_with_param(final, "download", "1"))
                if body[:2] == b"PK":
                    out("Downloaded workbook via %s, redirected (%d bytes)" % (name, len(body)))
                    return body, modified
            errors.append("%s: got a web page, not the Excel file" % name)
        except Exception as e:  # noqa: BLE001
            errors.append("%s: %s" % (name, e))
    raise RuntimeError("Couldn't download the workbook from OneDrive.\n  " + "\n  ".join(errors) +
                       "\nCheck that ONEDRIVE_URL is an 'Anyone with the link' share link.")


# ------------------------------------------------------------ data
def semester_overrides():
    raw = os.environ.get("SEMESTER_STARTS", "")
    res = {}
    for part in raw.replace("\n", ",").split(","):
        if "=" not in part:
            continue
        b, ym = part.split("=", 1)
        try:
            y, m = ym.strip().split("-")[:2]
            res[b.strip()] = (int(y), int(m))
        except ValueError:
            out("Ignoring SEMESTER_STARTS entry %r (use batch=YYYY-MM)" % part.strip())
    return res


def encrypt(plain, password):
    salt, iv = os.urandom(16), os.urandom(12)
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITERATIONS).derive(password.encode("utf-8"))
    ct = AESGCM(key).encrypt(iv, plain, None)
    b = lambda x: base64.b64encode(x).decode("ascii")  # noqa: E731
    return {"v": 1, "kdf": "PBKDF2-SHA256", "iter": ITERATIONS, "salt": b(salt), "iv": b(iv), "ct": b(ct)}


def hours_since_last_publish():
    try:
        t = subprocess.run(["git", "log", "-1", "--format=%ct", "--", "data.enc"], cwd=ROOT,
                           capture_output=True, text=True).stdout.strip()
        return (time.time() - int(t)) / 3600 if t else 1e9
    except Exception:  # noqa: BLE001
        return 1e9


def main():
    share_url = os.environ.get("ONEDRIVE_URL", "").strip()
    password = os.environ.get("DASHBOARD_PASSWORD", "")
    if not share_url or not password:
        out("Missing secret: add ONEDRIVE_URL and DASHBOARD_PASSWORD under Settings > Secrets and variables > Actions.")
        return 1

    body, last_modified = download(share_url)
    tmp = tempfile.mkdtemp(prefix="kw_")
    path = os.path.join(tmp, "Kingsway_Finance_Data.xlsx")
    with open(path, "wb") as f:
        f.write(body)

    data = kw_data._read_workbook(path, semester_overrides())
    when = None
    if last_modified:
        try:
            when = email.utils.parsedate_to_datetime(last_modified).astimezone(TZ)
        except Exception:  # noqa: BLE001
            when = None
    data["file_modified"] = when.strftime("%d %b %Y, %I:%M %p") if when else "OneDrive file"

    # Only publish when the figures changed, or every few hours as a heartbeat.
    fingerprint = hmac.new(password.encode("utf-8"),
                           json.dumps(data, sort_keys=True, default=str).encode("utf-8"),
                           hashlib.sha256).hexdigest()
    ver_path = os.path.join(ROOT, "data.ver")
    old = open(ver_path).read().strip() if os.path.exists(ver_path) else ""
    age = hours_since_last_publish()
    if fingerprint == old and age < FORCE_EVERY_H:
        out("No changes in the workbook; nothing to publish (last publish %.1f h ago)." % age)
        set_output("changed", "false")
        return 0

    data["warning"] = None
    data["user"] = "Kingsway"
    data["published"] = dt.datetime.now(TZ).isoformat(timespec="seconds")
    data["workbook_name"] = "Kingsway_Finance_Data.xlsx"
    plain = json.dumps(data, separators=(",", ":"), default=str).encode("utf-8")
    with open(os.path.join(ROOT, "data.enc"), "w", encoding="utf-8") as f:
        json.dump(encrypt(plain, password), f)
    with open(ver_path, "w") as f:
        f.write(fingerprint + "\n")
    out("Published new figures (%d bytes)." % len(plain))
    set_output("changed", "true")
    return 0


if __name__ == "__main__":
    sys.exit(main())
