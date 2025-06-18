import argparse
import os
import sys
import json
from pathlib import Path

def create_windows_shortcut(command_name, icon_override=None):
    import pythoncom
    from win32com.client import Dispatch
    import cptd_tools

    desktop = Path(os.environ["USERPROFILE"]) / "Desktop"
    shortcut_path = desktop / f"{command_name}.lnk"

    # The way to cptd.exe
    scripts_dir = Path(sys.executable).parent / "Scripts"
    cptd_exe = scripts_dir / "cptd.exe"
    if not cptd_exe.exists():
        print(f"[!] Not found cptd.exe в {scripts_dir}")
        return

    # The way to command directory
    base_dir = Path(cptd_tools.__file__).parent
    command_path = base_dir / "commands" / command_name

    if not command_path.exists():
        print(f"[!] Team folder not found: {command_path}")
        return

    # Reading the manifesto
    manifest_path = command_path / "manifest.json"
    if not manifest_path.exists():
        print(f"[!] Not found manifest.json: {manifest_path}")
        return

    with open(manifest_path, "r", encoding="utf-8") as f:
        manifest = json.load(f)

    entrypoint = manifest.get("entrypoint")
    icon_declared = manifest.get("icon")

    if not entrypoint:
        print(f"[!] Not specified in the manifest entrypoint")
        return

    entrypoint_path = command_path / entrypoint
    if not entrypoint_path.exists():
        print(f"[!] File entrypoint not found: {entrypoint_path}")
        return

    # Search for icon
    icon_path = None
    if icon_override:
        icon_candidate = Path(icon_override).expanduser().resolve()
        if icon_candidate.exists():
            icon_path = icon_candidate
        else:
            print(f"[!] The specified icon was not found.: {icon_candidate}")
    elif icon_declared:
        declared = Path(icon_declared)
        if declared.suffix:
            icon_candidate = entrypoint_path.parent / declared
            if icon_candidate.exists():
                icon_path = icon_candidate.resolve()
            else:
                alt_ext = ".ico" if declared.suffix.lower() == ".png" else ".png"
                alt_candidate = icon_candidate.with_suffix(alt_ext)
                if alt_candidate.exists():
                    icon_path = alt_candidate.resolve()
                    print(f"[ℹ] Icon found by alternative extension: {alt_candidate}")
                else:
                    print(f"[ℹ] Icon not found: {icon_candidate} or {alt_candidate}")
        else:
            for ext in [".ico", ".png"]:
                candidate = entrypoint_path.parent / (declared.name + ext)
                if candidate.exists():
                    icon_path = candidate.resolve()
                    break

    # Create a shortcut
    shell = Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(str(shortcut_path))
    shortcut.TargetPath = str(cptd_exe.resolve())
    shortcut.Arguments = command_name
    shortcut.WorkingDirectory = str(desktop)
    if icon_path:
        shortcut.IconLocation = str(icon_path)
    shortcut.save()

    print(f"[✓] Shortcut created: {shortcut_path}")
    if icon_path:
        print(f"[✓] The icon is installed: {icon_path}")
    else:
        print("[ℹ] Icon not specified or not found - default is used.")

def delete_windows_shortcut(command_name):
    shortcut_path = Path(os.environ["USERPROFILE"]) / "Desktop" / f"{command_name}.lnk"
    if shortcut_path.exists():
        shortcut_path.unlink()
        print(f"[✓] Shortcut removed: {shortcut_path}")
    else:
        print(f"[ℹ] Label not found: {shortcut_path}")

def run(argv):
    parser = argparse.ArgumentParser(description="Adding or removing shortcuts CPTD")
    parser.add_argument("--add", help="Command name to add a shortcut")
    parser.add_argument("--delete", help="Command name to remove shortcut")
    parser.add_argument("--icon", help="Path to icon (.ico or .png)", required=False)
    args = parser.parse_args(argv)

    if args.add:
        create_windows_shortcut(args.add, args.icon)
    elif args.delete:
        delete_windows_shortcut(args.delete)
    else:
        print("[ℹ] Use it --add or--delete with the team name.")
