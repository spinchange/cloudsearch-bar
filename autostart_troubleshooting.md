# CloudSearchBar Autostart Troubleshooting

## 1. Verify the registry entry was written
Open `regedit` and navigate to:
```
HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run
```
Look for a `CloudSearchBar` entry. If it's missing, the program either never ran,
or `autostart` was `false` in the ini when it launched.

## 2. If the entry is missing — run the exe once
The registry is only written when the program starts. Launch the exe once
(with `autostart = true` in the ini) to register it, then restart and test.

## 3. Check the path in the registry entry
If the entry exists, confirm the path it points to is correct and the exe
hasn't been moved or renamed since it was first run.

## 4. Check the ini on the work computer
Open `cloudsearch_bar.ini` next to the exe and confirm:
```ini
[Startup]
autostart = true
```

## 5. Work computer security / IT policy
Corporate machines often have Group Policy that blocks HKCU\Run entries or
kills unknown processes on startup. Check with IT or look in Event Viewer
(eventvwr) -> Windows Logs -> Application for any blocked/killed entries.

## 6. Antivirus quarantine
The exe might launch and immediately get killed by AV. Check your AV's
quarantine or block log.

## 7. Tray icon is hidden
It might actually be running but the tray icon is tucked in the overflow area
(the ^ arrow in the taskbar). Check there before assuming it didn't start.

## Most likely causes on a work machine
- Group Policy blocking HKCU\Run entries (#5)
- Exe was never run to write the registry entry (#2)
