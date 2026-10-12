# Windows app launcher thingy
# Note: I wrote this because the standard Windows search is too slow sometimes. 
# It goes through 4 tiers to find whatever you type.
import csv
import difflib
import glob
import os
import shutil
import subprocess

# Global vars for AppOpener, kinda hate globals but whatever, it works
ao_open = None
ao_names = None

def _ensure_appopener():
    global ao_open, ao_names
    if ao_open is not None:
        return True
    
    try:
        from AppOpener import open as _o, give_appnames as _g
    except ImportError:
        # User probably didn't pip install it, oh well
        return False
        
    ao_open = _o
    ao_names = _g
    return True

class AppLauncher:
    def __init__(self):
        # Gotta be on windows obviously
        if os.name != "nt":
            raise SystemExit("Windows only, sorry!")
            
        self._shell = None
        self._uwp = None
        self._lnks = None

    # Tier 1: Check if it's straight up on the PATH
    def _tier1(self, app):
        p = shutil.which(app)
        if p:
            subprocess.Popen([p])
            return f"Opened {app} via PATH."
        return None

    # Tier 2: AppOpener fuzzy match index
    def _tier2(self, app):
        if not _ensure_appopener():
            return None
            
        names = sorted(ao_names())
        q = app.lower()
        
        # Doing a bunch of checks to find the best name match
        best = next((n for n in names if q == n.lower()), None) or \
            next((n for n in names if q in n.lower()), None) or \
            next(iter(difflib.get_close_matches(app, names, n=1, cutoff=0.6)), None)
            
        if best is None:
            return None
            
        ao_open(best, output=False)
        return f"Opened {best} via AppOpener."

    # Tier 3: PowerShell Get-StartApps for UWP / Windows Store apps
    def _tier3(self, app):
        if self._uwp is None:
            # Let's run powershell to grab store apps. Might be a bit slow on first call.
            r = subprocess.run(
                ["powershell", "-NoProfile", "-Command",
                 "Get-StartApps | Select-Object Name,AppID | ConvertTo-Csv -NoTypeInformation"],
                capture_output=True, text=True)
            
            if r.returncode == 0:
                self._uwp = list(csv.DictReader(r.stdout.splitlines()))
            else:
                self._uwp = []
                
        q = app.lower()
        hit = next((a for a in self._uwp if q in (a.get("Name") or "").lower()), None)
        
        if hit is None and self._uwp:
            m = difflib.get_close_matches(app, [a.get("Name") or "" for a in self._uwp], n=1, cutoff=0.6)
            hit = next((a for a in self._uwp if a.get("Name") == (m[0] if m else None)), None)
            
        if hit:
            subprocess.Popen(["explorer", f"shell:appsFolder\\{hit['AppID']}"])
            return f"Opened {hit['Name']} via Store."
            
        return None

    # Tier 4: Digging through the Start Menu .lnk shortcuts manually
    def _tier4(self, app):
        if self._lnks is None:
            self._lnks = {}
            try:
                import win32com.client
                if self._shell is None:
                    self._shell = win32com.client.Dispatch("WScript.Shell")
                    
                # Look in both common and user start menus
                program_data = os.environ.get("ProgramData", "")
                app_data = os.environ.get("AppData", "")
                
                paths_to_check = [
                    os.path.join(program_data, r"Microsoft\Windows\Start Menu\Programs"),
                    os.path.join(app_data, r"Microsoft\Windows\Start Menu\Programs")
                ]
                
                for base in paths_to_check:
                    if os.path.exists(base):
                        for lnk in glob.glob(os.path.join(base, "**", "*.lnk"), recursive=True):
                            key_name = os.path.splitext(os.path.basename(lnk))[0].lower()
                            self._lnks[key_name] = lnk
            except Exception as e:
                # print(e) # debug later maybe
                return None
                
        q = app.lower()
        key = next((k for k in self._lnks if q in k), None)
        
        if key is None and self._lnks:
            m = difflib.get_close_matches(q, list(self._lnks), n=1, cutoff=0.6)
            key = m[0] if m else None
            
        if key:
            os.startfile(self._lnks[key])
            return f"Opened {key} via shortcut."
            
        return None

    def open(self, app_name: str) -> str:
        # Loop through all 4 tiers until something works
        tiers = [self._tier1, self._tier2, self._tier3, self._tier4]
        
        for t in tiers:
            try:
                hit = t(app_name.strip())
            except Exception:
                hit = None
                
            if hit:
                return hit
                
        return f"Oops, couldn't find '{app_name}' anywhere."

def main():
    print("Initializing launcher... please wait a sec.")
    launcher = AppLauncher()
    
    # Try to load names for autocomplete/suggestions if AppOpener is installed
    if _ensure_appopener():
        names = sorted(ao_names())
    else:
        names = []
        
    print(f"Ready! Indexed {len(names)} apps. Type a name, use commas for multiple, or 'q' to exit.")
    
    while True:
        try:
            q = input("launch> ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nBye!")
            break
            
        if q.lower() in ("q", "quit", "exit"):
            break
            
        if q.lower() == "refresh":
            launcher._uwp = None
            launcher._lnks = None
            if _ensure_appopener():
                names = sorted(ao_names())
            print(f"Cache cleared and refreshed. ({len(names)} apps)")
            continue
            
        if not q:
            continue
            
        # Support launching multiple apps separated by comma
        if "," in q:
            results = []
            for part in q.split(","):
                if part.strip():
                    results.append(launcher.open(part.strip()))
            print("; ".join(results))
            continue
            
        # If it's not a direct hit, let's show a few suggestions just in case
        if not q.isdigit() and len(names) > 0:
            hits = [n for n in names if q.lower() in n.lower()][:10]
            if not hits:
                hits = difflib.get_close_matches(q, names, n=10, cutoff=0.6)
                
            if hits and q.lower() not in (h.lower() for h in hits):
                print("Did you mean one of these?")
                for i, h in enumerate(hits, 1):
                    print(f"  {i}. {h}")
                    
                choice = input("Select number or press enter to search normally: ").strip()
                if choice.isdigit():
                    idx = int(choice) - 1
                    if 0 <= idx < len(hits):
                        q = hits[idx]
                        
        print(launcher.open(q))

if __name__ == "__main__":
    main()