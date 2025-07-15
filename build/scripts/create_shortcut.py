import os
import sys
import subprocess
from pathlib import Path

def create_shortcut():
    desktop = Path(os.path.expanduser("~/Desktop"))
    target = Path(sys.executable).parent / "outlook_extractor.exe"
    shortcut = desktop / "Outlook Extractor.lnk"
    
    vbs_script = f"""
    Set oWS = WScript.CreateObject("WScript.Shell")
    sLinkFile = "{shortcut}"
    Set oLink = oWS.CreateShortcut(sLinkFile)
    oLink.TargetPath = "{target}"
    oLink.WorkingDirectory = "{target.parent}"
    oLink.Save
    """
    
    vbs_path = Path(__file__).parent / "create_shortcut.vbs"
    with open(vbs_path, "w") as f:
        f.write(vbs_script)
    
    subprocess.call(['cscript.exe', str(vbs_path)])
    vbs_path.unlink()

if __name__ == "__main__":
    create_shortcut()
