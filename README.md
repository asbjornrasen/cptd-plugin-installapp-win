# cptd-plugin-installapp-win
---

## installapp Command (Windows)

The `installapp` command creates or removes shortcuts for installed CPTD commands, allowing them to be launched as standard desktop applications in Windows.

---

### Usage

```
cptd installapp --add <command>
cptd installapp --delete <command> [--icon <path>]
```

---

### Arguments

| Argument   | Required            | Description                                                |
| ---------- | ------------------- | ---------------------------------------------------------- |
| `--add`    | yes (or `--delete`) | The name of the installed CPTD command to add as shortcut  |
| `--delete` | yes (or `--add`)    | The name of the command to remove the shortcut             |
| `--icon`   | no                  | Path to an optional `.ico` or `.png` icon for the shortcut |

---

### Examples

```
cptd installapp --add testapp
cptd installapp --add testapp --icon C:\Icons\custom.ico
cptd installapp --delete testapp
```

---

### Behavior

* Automatically creates a `.lnk` shortcut that launches `cptd <command>` via terminal.
* Uses the provided icon, or falls back to the one in the manifest or a default icon.
* Removes the corresponding shortcut when `--delete` is used.

---

### Shortcut Location

Shortcuts are automatically added to:

```
%USERPROFILE%\Desktop
%APPDATA%\Microsoft\Windows\Start Menu\Programs
```

No manual move is required.

---

### Requirements

* Windows 10 or higher
* Python installed and CPTD available in system PATH
* Command is executed in `cmd.exe` or `PowerShell`

---
