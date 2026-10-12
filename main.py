"""Windows-only TUI app opener: 4-tier cascade, fuzzy search UI, no AI."""
import csv
import difflib  # ponytail: last-resort typo net only; token matcher ranks first
import glob
import os
import re
import shutil
import subprocess

from prompt_toolkit import prompt as pt_prompt
from prompt_toolkit.completion import Completer, Completion

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


def _tokens(s):
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).split()


def _match(query, names, limit=10):
    """Rank names: exact > word-prefix > initialism > substring > token-subset > word-typo > typo."""
    qt = _tokens(query)
    qraw = " ".join(qt)
    if not qraw:
        return []
    # ponytail: tokenize once per call so per-keystroke ranking stays fast
    seq = list(names)
    toks = seq if seq and isinstance(seq[0], tuple) else [(n, _tokens(n)) for n in seq]
    scored = []
    for n, nt in toks:
        if not nt:
            continue
        nraw = " ".join(nt)
        if qraw == nraw:
            s = 100
        elif all(any(w.startswith(t) for w in nt) for t in qt):
            s = 80
        elif _initialism_hit(qt, nt):
            s = 70
        elif qraw in nraw:
            s = 60
        elif all(t in nraw for t in qt):
            s = 50
        elif all(any(w.startswith(t) or difflib.SequenceMatcher(None, t, w).ratio() >= 0.8
                     for w in nt) for t in qt):
            s = 30
        else:
            continue
        scored.append((s, n))
    scored.sort(key=lambda x: (-x[0], len(x[1])))
    return [n for _, n in scored[:limit]]


def _initialism_hit(qt, nt):
    initials = "".join(w[0] for w in nt)
    return all(any(w.startswith(t) for w in nt) or initials.startswith(t) for t in qt)


class AppLauncher:
    def __init__(self):
        if os.name != "nt":
            raise SystemExit("Windows only.")
        self._shell = None
        self._uwp = None
        self._lnks = None
        self._exes = None
        self._index = None
        self._tok = None

    def _uwp_apps(self):
        if self._uwp is None:
            r = subprocess.run(
                ["powershell", "-NoProfile", "-Command",
                 "Get-StartApps | Select-Object Name,AppID | ConvertTo-Csv -NoTypeInformation"],
                capture_output=True, text=True)
            self._uwp = list(csv.DictReader(r.stdout.splitlines())) if r.returncode == 0 else []
        return self._uwp

    def _shortcuts(self):
        if self._lnks is None:
            self._lnks = {}
            try:
                import win32com.client
                if self._shell is None:
                    self._shell = win32com.client.Dispatch("WScript.Shell")
                for base in (os.path.join(os.environ.get("ProgramData", ""), r"Microsoft\Windows\Start Menu\Programs"),
                             os.path.join(os.environ.get("AppData", ""), r"Microsoft\Windows\Start Menu\Programs")):
                    for lnk in glob.glob(os.path.join(base, "**", "*.lnk"), recursive=True):
                        self._lnks.setdefault(os.path.splitext(os.path.basename(lnk))[0], lnk)
            except Exception:
                pass
        return self._lnks

    def _path_exes(self):
        if self._exes is None:
            self._exes = {}
            seen = set()
            for d in os.environ.get("PATH", "").split(os.pathsep):
                d = d.strip('"')
                if not d or d.lower() in seen or not os.path.isdir(d):
                    continue
                seen.add(d.lower())
                try:
                    for f in os.listdir(d):
                        if f.lower().endswith(".exe"):
                            self._exes.setdefault(os.path.splitext(f)[0], os.path.join(d, f))
                except OSError:
                    pass
        return self._exes

    def names(self):
        if self._index is None:
            idx = {}
            if _ensure_appopener():
                for n in ao_names():
                    idx.setdefault(n, n)
            for a in self._uwp_apps():
                if a.get("Name"):
                    idx.setdefault(a["Name"], a["Name"])
            for n in self._shortcuts():
                idx.setdefault(n, n)
            for n in self._path_exes():
                idx.setdefault(n, n)
            self._index = sorted(idx, key=str.lower)
        return self._index

    def match(self, query, limit=10):
        if self._tok is None:
            self._tok = [(n, _tokens(n)) for n in self.names()]
        return _match(query, self._tok, limit)

    def rescan(self):
        try:
            from AppOpener import update_list
            update_list.update(output=False)
        except Exception:
            pass
        self._uwp = self._lnks = self._exes = self._index = self._tok = None
        return len(self.names())

    def _tier1(self, app):
        p = shutil.which(app)
        if p:
            subprocess.Popen([p])
            return f"Opened {app} via PATH."
        hit = _match(app, self._path_exes(), 1)
        if hit:
            subprocess.Popen([self._path_exes()[hit[0]]])
            return f"Opened {hit[0]} via PATH."
        return None

    def _tier2(self, app):
        if not _ensure_appopener():
            return None
        hit = _match(app, ao_names(), 1)
        if not hit:
            return None
        ao_open(hit[0], output=False)
        return f"Opened {hit[0]} via AppOpener."

    def _tier3(self, app):
        uwps = self._uwp_apps()
        hit = _match(app, [a.get("Name") or "" for a in uwps], 1)
        if not hit:
            return None
        appid = next(a["AppID"] for a in uwps if a.get("Name") == hit[0])
        subprocess.Popen(["explorer", f"shell:appsFolder\\{appid}"])
        return f"Opened {hit[0]} via Store."

    def _tier4(self, app):
        lnks = self._shortcuts()
        hit = _match(app, lnks, 1)
        if not hit:
            return None
        os.startfile(lnks[hit[0]])  # noqa: S606 (local trusted shortcut)
        return f"Opened {hit[0]} via shortcut."

    def open(self, app_name: str) -> str:
        app_name = app_name.strip()
        for tier in (self._tier1, self._tier2, self._tier3, self._tier4):
            try:
                hit = tier(app_name)
            except Exception:
                hit = None
            if hit:
                return hit
        return f"'{app_name}' not found."


class _Completer(Completer):
    def __init__(self, launcher):
        self._launcher = launcher

    def get_completions(self, doc, _event):
        text = doc.text_before_cursor
        if not text.strip():
            return
        for n in self._launcher.match(text, 12):
            yield Completion(n, start_position=-len(text))


def main():
    os.system("cls")
    print("Loading... (first run scans apps, once)")
    launcher = AppLauncher()
    print(f"{len(launcher.names())} apps indexed. Type to search, arrows to pick, q to quit.")
    completer = _Completer(launcher)
    use_pt, last_hits = True, []
    while True:
        try:
            if use_pt:
                try:
                    q = pt_prompt("launch> ", completer=completer, complete_while_typing=True).strip()
                except Exception:
                    use_pt = False  # ponytail: non-console host (pipes), plain input still works
                    continue
            else:
                q = input("launch> ").strip()
                if q.isdigit() and last_hits:
                    q = last_hits[int(q) - 1] if 1 <= int(q) <= len(last_hits) else q
                last_hits = []
        except (EOFError, KeyboardInterrupt):
            break
        if q.lower() in ("q", "quit", "exit"):
            break
        if q.lower() in ("rescan", "refresh"):
            print(f"Rescanned ({launcher.rescan()} apps).")
            continue
        if not q:
            continue
        if "," in q:
            print("; ".join(launcher.open(p) for p in q.split(",") if p.strip()))
            continue
        if not use_pt:
            hits = launcher.match(q, 10)
            if hits and q.lower() not in (h.lower() for h in hits):
                for i, h in enumerate(hits, 1):
                    print(f"{i}. {h}")
                last_hits = hits
                continue
        print(launcher.open(q))


if __name__ == "__main__":
    main()
