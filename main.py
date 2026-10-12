#!/usr/bin/env python
# -*- coding: utf-8 -*-
# app launcher thingy. windows only cause thats what i use
# type something, it finds the app, it opens. thats the whole idea
# started befcause windows search kept losing my apps. 1300+ on this pc and counting
import csv
import difflib
import glob
import os
import re
import shutil
import subprocess

from prompt_toolkit import print_formatted_text, prompt as pt_prompt
from prompt_toolkit.completion import Completer, Completion
from prompt_toolkit.formatted_text import HTML

ao_open = None
ao_names = None

# picked prompt_toolkit over textual cause textual pulls half of pypi for a launcher.
# stdlib input() was fine until i wanted the arrow-key dropdown. so here we are
# programdata never moves on win10/11. appdata obviously depends on the user
STARTMENUS = [
    r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs",
    os.path.join(os.environ.get("AppData", ""), r"Microsoft\Windows\Start Menu\Programs"),
]

# colors per tier, looks nice in windows terminal. cmd.exe kinda washes them out but eh
TIER_TAG = {1: ("PATH", "ansigreen"), 2: ("AppOpener", "ansicyan"),
            3: ("Store", "ansimagenta"), 4: ("shortcut", "ansiyellow")}


def say(text, tier=0):
    try:
        if tier in TIER_TAG:
            tag, color = TIER_TAG[tier]
            print_formatted_text(HTML(f"<{color}>[{tag}]</{color}> {text}"))
        else:
            # red for misses, plain otherwise
            print_formatted_text(HTML(f"<ansired>{text}</ansired>") if tier == -1 else text)
    except Exception:
        print(text)  # pipes and stuff, whatever


def _ensure_appopener():
    global ao_open, ao_names
    if ao_open is not None:
        return True
    try:
        from AppOpener import open as _o, give_appnames as _g
    except ImportError:
        return False  # pip install AppOpener if you want tier 2
    ao_open, ao_names = _o, _g
    return True


def _tokens(s):
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).split()

def _rank(query, names, limit=10):
    # first version was just difflib.get_close_matches over everything.
    # I typed "vs code", got discord. never again. token ranking now - exact,
    # prefix, initials, substring, all-words-there, typo-per-word, in that order
    qt = _tokens(query)
    qraw = " ".join(qt)
    if not qraw:
        return []
    seq = list(names)
    toks = seq if seq and isinstance(seq[0], tuple) else [(n, _tokens(n)) for n in seq]
    scored = []
    for n, nt in toks:
        if not nt:
            continue
        nraw = " ".join(nt)
        if qraw == nraw:
            s = 100
        else:
            initials = "".join(w[0] for w in nt)  # "vs" matching visual-studio. vscode never matched without this lol
            if all(any(w.startswith(t) for w in nt) for t in qt):
                s = 80
            elif all(any(w.startswith(t) for w in nt) or initials.startswith(t) for t in qt):
                s = 70
            elif qraw in nraw:
                s = 60
            elif all(t in nraw for t in qt):
                s = 50
            elif all(any(w.startswith(t) or difflib.SequenceMatcher(None, t, w).ratio() >= 0.8
                         for w in nt) for t in qt):
                s = 30  # typos. chorme -> chrome lands here. 0.8 not 0.75: at 0.75 "app" matched "pp"
                        # (ratio 0.80) and it launched pp-trace for a garbage query. learned that the fun way
            else:
                continue
        # print(s, n)  # keeping this for next time i debug the ranking
        scored.append((s, n))
    scored.sort(key=lambda x: (-x[0], len(x[1])))
    return [n for _, n in scored[:limit]]


class AppLauncher:
    def __init__(self):
        if os.name != "nt":
            raise SystemExit('Windows only, sorry') # I wanted to write LOL, but professionalism does not allow that :( 
        self._shell = None
        self._uwp = None
        self._lnks = None
        self._exes = None
        self._index = None
        self._tok = None

    def _uwp_apps(self):
        if self._uwp is None:
            # powershell is slow-ish the first time, then its cached so who cares
            r = subprocess.run(
                ["powershell", "-NoProfile", "-Command",
                 "Get-StartApps | Select-Object Name,AppID | ConvertTo-Csv -NoTypeInformation"],
                capture_output=True, text=True)
            if r.returncode != 0:
                say("store lookup broke (ps policy stuff?), skipping those apps this time")
                self._uwp = []
            else:
                self._uwp = list(csv.DictReader(r.stdout.splitlines()))
        return self._uwp

    def _shortcuts(self):
        if self._lnks is None:
            self._lnks = {}
            try:
                import win32com.client
                if self._shell is None:
                    self._shell = win32com.client.Dispatch("WScript.Shell")
                for base in STARTMENUS:  # user one usually has the good stuff
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
                    continue  # locked system folders etc, just skip
        return self._exes

    def names(self):
        # one big dict of everything. first one wins if names clash
        if self._index is None:
            idx = {}
            if _ensure_appopener():
                for n in ao_names():
                    idx.setdefault(n, 2)
            for a in self._uwp_apps():
                if a.get("Name"):
                    idx.setdefault(a["Name"], 3)
            for n in self._shortcuts():
                idx.setdefault(n, 4)
            for n in self._path_exes():
                idx.setdefault(n, 1)
            self._index = dict(sorted(idx.items(), key=lambda kv: kv[0].lower()))
        return self._index

    def tier_of(self, name):
        return self.names().get(name, 0)

    def match(self, query, limit=10):
        if self._tok is None:
            self._tok = [(n, _tokens(n)) for n in self.names()]
        return _rank(query, self._tok, limit)

    def rescan(self):
        try:
            from AppOpener import update_list
            update_list.update(output=False)
        except Exception:
            pass
        self._uwp = self._lnks = self._exes = self._index = self._tok = None
        return len(self.names())

    # TODO: tier 1 should probably check .bat and .cmd too, whatever, exes cover 99%
    def _tier1(self, app):
        p = shutil.which(app)
        if p:
            subprocess.Popen([p])
            return f"Opened {app} via PATH.", 1
        hit = _rank(app, self._path_exes(), 1)
        if hit:
            subprocess.Popen([self._path_exes()[hit[0]]])
            return f"Opened {hit[0]} via PATH.", 1
        return None

    def _tier2(self, app):
        if not _ensure_appopener():
            return None
        hit = _rank(app, ao_names(), 1)
        if not hit:
            return None
        ao_open(hit[0], output=False)
        return f"Opened {hit[0]} via AppOpener.", 2

    def _tier3(self, app):
        uwps = self._uwp_apps()
        found = _rank(app, [a.get("Name") or "" for a in uwps], 1)
        if not found:
            return None
        appid = next(a["AppID"] for a in uwps if a.get("Name") == found[0])
        subprocess.Popen(["explorer", f"shell:appsFolder\\{appid}"])
        return f"Opened {found[0]} via Store.", 3

    def _tier4(self, app):
        lnks = self._shortcuts()
        res = _rank(app, lnks, 1)
        if not res:
            return None
        os.startfile(lnks[res[0]])
        return f"Opened {res[0]} via shortcut.", 4

    def open(self, app_name: str):
        app_name = app_name.strip()
        for tier in (self._tier1, self._tier2, self._tier3, self._tier4):
            try:
                hit = tier(app_name)
            except Exception:
                hit = None
            if hit:
                return hit
        return f"'{app_name}' not found.", -1


class _Completer(Completer):
    def __init__(self, launcher):
        self._launcher = launcher

    def get_completions(self, doc, _event):
        text = doc.text_before_cursor
        if not text.strip():
            return
        for n in self._launcher.match(text, 12):
            tag = TIER_TAG.get(self._launcher.tier_of(n), ("", ""))[0]
            yield Completion(n, start_position=-len(text), display_meta=tag or None)


HELP = """Commands:
  <typing>          live search, Tab completes, arrows pick, enter launches
  name, name, ...  launch several apps at once
  rescan            re-fetch newly installed apps
  help              this text
  q                 quit
  eg. type "vs code" and it actually finds vscode now"""

# I generated it from a site called patorjk.com hehe >~<
First_Statement = """
 ██╗ ██████╗  ██████╗  ██████╗ ██╗ ██╗    ██████╗ ███████╗██╗     ██╗ █████╗ ██████╗ ██╗     ███████╗
███║██╔═████╗██╔═████╗██╔═████╗╚═╝██╔╝    ██╔══██╗██╔════╝██║     ██║██╔══██╗██╔══██╗██║     ██╔════╝
╚██║██║██╔██║██║██╔██║██║██╔██║  ██╔╝     ██████╔╝█████╗  ██║     ██║███████║██████╔╝██║     █████╗  
 ██║████╔╝██║████╔╝██║████╔╝██║ ██╔╝      ██╔══██╗██╔══╝  ██║     ██║██╔══██║██╔══██╗██║     ██╔══╝  
 ██║╚██████╔╝╚██████╔╝╚██████╔╝██╔╝██╗    ██║  ██║███████╗███████╗██║██║  ██║██████╔╝███████╗███████╗
 ╚═╝ ╚═════╝  ╚═════╝  ╚═════╝ ╚═╝ ╚═╝    ╚═╝  ╚═╝╚══════╝╚══════╝╚═╝╚═╝  ╚═╝╚═════╝ ╚══════╝╚══════╝
                                                                                                     
 █████╗ ██████╗ ██████╗     ██████╗ ██████╗ ███████╗███╗   ██╗███████╗██████╗                        
██╔══██╗██╔══██╗██╔══██╗   ██╔═══██╗██╔══██╗██╔════╝████╗  ██║██╔════╝██╔══██╗                       
███████║██████╔╝██████╔╝   ██║   ██║██████╔╝█████╗  ██╔██╗ ██║█████╗  ██████╔╝                       
██╔══██║██╔═══╝ ██╔═══╝    ██║   ██║██╔═══╝ ██╔══╝  ██║╚██╗██║██╔══╝  ██╔══██╗                       
██║  ██║██║     ██║███████╗╚██████╔╝██║     ███████╗██║ ╚████║███████╗██║  ██║                       
╚═╝  ╚═╝╚═╝     ╚═╝╚══════╝ ╚═════╝ ╚═╝     ╚══════╝╚═╝  ╚═══╝╚══════╝╚═╝  ╚═╝                       
                                                                                                     
"""


def main():
    os.system("cls")
    say(First_Statement)
    say("Loading... (first run scans apps, once)")
    launcher = AppLauncher()
    n = len(launcher.names())
    say("=" * 46)
    say("  APP LAUNCHER - {} apps indexed, type to search".format(n))
    say("  arrows pick - enter launches - 'help' for commands")
    say("=" * 46)
    completer = _Completer(launcher)
    use_pt, last_hits = True, []
    while True:
        try:
            if use_pt:
                try:
                    q = pt_prompt(HTML("<ansicyan>launch></ansicyan> "),
                                  completer=completer, complete_while_typing=True).strip()
                except Exception:
                    use_pt = False  # no console (pipes etc), fall back to boring input
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
            say(f"Rescanned ({launcher.rescan()} apps).")
            continue
        if q.lower() == "help":
            say(HELP)
            continue
        if not q:
            continue
        if "," in q:
            for p in q.split(","):
                if p.strip():
                    msg, tier = launcher.open(p)
                    say(msg, tier)
            continue
        if not use_pt:
            hits = launcher.match(q, 10)
            if hits and q.lower() not in (h.lower() for h in hits):
                for i, h in enumerate(hits, 1):
                    tag = TIER_TAG.get(launcher.tier_of(h), ("", ""))[0]
                    say(f"{i}. {h}" + (f" [{tag}]" if tag else ""))
                last_hits = hits
                continue
        msg, tier = launcher.open(q)
        say(msg, tier)



#psst...just testing it
if __name__ == "__main__":
    main()
