"""Windows-only TUI app opener: 4-tier cascade, stdlib UI, no AI."""
import csv
import difflib
import glob
import os
import shutil
import subprocess

# ponytail: `open` shadows builtin, so AppOpener's is lazily aliased as ao_open
ao_open = None
ao_names = None


def _ensure_appopener():
    global ao_open, ao_names
    if ao_open is not None:
        return True
    try:
        from AppOpener import open as _o, give_appnames as _g
    except ImportError:
        return False
    ao_open, ao_names = _o, _g
    return True


class AppLauncher:
    def __init__(self):
        if os.name != "nt":
            raise SystemExit("Windows only.")
        self._shell = None
        self._uwp = None
        self._lnks = None

    # Tier 1: executables on PATH
    def _tier1(self, app):
        p = shutil.which(app)
        if p:
            subprocess.Popen([p])
            return f"Opened {app} via PATH."
        return None

    # Tier 2: AppOpener fuzzy index
    # ponytail: throw_error only fires when output=True upstream, so pre-match here and keep output off
    def _tier2(self, app):
        if not _ensure_appopener():
            return None
        names = sorted(ao_names())
        q = app.lower()
        best = next((n for n in names if q == n.lower()), None) or \
            next((n for n in names if q in n.lower()), None) or \
            next(iter(difflib.get_close_matches(app, names, n=1, cutoff=0.6)), None)
        if best is None:
            return None
        ao_open(best, output=False)
        return f"Opened {best} via AppOpener."

    # Tier 3: live UWP / Store lookup (same source AppOpener indexes)
    def _tier3(self, app):
        if self._uwp is None:
            r = subprocess.run(
                ["powershell", "-NoProfile", "-Command",
                 "Get-StartApps | Select-Object Name,AppID | ConvertTo-Csv -NoTypeInformation"],
                capture_output=True, text=True)
            self._uwp = list(csv.DictReader(r.stdout.splitlines())) if r.returncode == 0 else []
        q = app.lower()
        hit = next((a for a in self._uwp if q in (a.get("Name") or "").lower()), None)
        if hit is None and self._uwp:
            m = difflib.get_close_matches(app, [a.get("Name") or "" for a in self._uwp], n=1, cutoff=0.6)
            hit = next((a for a in self._uwp if a.get("Name") == (m[0] if m else None)), None)
        if hit:
            subprocess.Popen(["explorer", f"shell:appsFolder\\{hit['AppID']}"])
            return f"Opened {hit['Name']} via Store."
        return None

    # Tier 4: Start Menu .lnk deep scan
    def _tier4(self, app):
        if self._lnks is None:
            self._lnks = {}
            try:
                import win32com.client
                if self._shell is None:
                    self._shell = win32com.client.Dispatch("WScript.Shell")
                for base in (os.path.join(os.environ.get("ProgramData", ""), r"Microsoft\Windows\Start Menu\Programs"),
                             os.path.join(os.environ.get("AppData", ""), r"Microsoft\Windows\Start Menu\Programs")):
                    for lnk in glob.glob(os.path.join(base, "**", "*.lnk"), recursive=True):
                        self._lnks[os.path.splitext(os.path.basename(lnk))[0].lower()] = lnk
            except Exception:
                return None
        q = app.lower()
        key = next((k for k in self._lnks if q in k), None)
        if key is None and self._lnks:
            m = difflib.get_close_matches(q, list(self._lnks), n=1, cutoff=0.6)
            key = m[0] if m else None
        if key:
            os.startfile(self._lnks[key])  # noqa: S606 (local trusted shortcut)
            return f"Opened {key} via shortcut."
        return None

    def open(self, app_name: str) -> str:
        for tier in (self._tier1, self._tier2, self._tier3, self._tier4):
            try:
                hit = tier(app_name.strip())
            except Exception:
                hit = None
            if hit:
                return hit
        return f"{app_name} not found."


def main():
    print("Loading... (first run scans apps, once)")
    launcher = AppLauncher()
    names = sorted(ao_names()) if _ensure_appopener() else []
    print(f"{len(names)} apps indexed. Type to filter, number to open, q to quit.")
    while True:
        try:
            q = input("> ").strip()
        except (EOFError, KeyboardInterrupt):
            break
        if q.lower() in ("q", "quit", "exit"):
            break
        if q.lower() == "refresh":
            launcher._uwp = launcher._lnks = None
            names = sorted(ao_names()) if _ensure_appopener() else names
            print(f"Refreshed ({len(names)} apps).")
            continue
        if not q:
            continue
        if "," in q:  # ponytail: AppOpener supports multi-launch, so do we
            print("; ".join(launcher.open(p) for p in q.split(",") if p.strip()))
            continue
        if not q.isdigit():
            hits = [n for n in names if q.lower() in n.lower()][:10] or \
                difflib.get_close_matches(q, names, n=10, cutoff=0.6)
            if hits and q.lower() not in (h.lower() for h in hits):
                for i, h in enumerate(hits, 1):
                    print(f"{i}. {h}")
                q = input("# or name> ").strip()
                if q.isdigit():
                    q = hits[int(q) - 1] if 1 <= int(q) <= len(hits) else q
        print(launcher.open(q))


if __name__ == "__main__":
    main()
