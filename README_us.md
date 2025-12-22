# Tray Icon Tool (TrayIcon)

This is a VB.NET-based Windows tray application. It displays a tray icon on startup and dynamically builds menu items from CSV files.

## Features

- Load menu structure from CSV files
- Supports 3-level hierarchy: Main Menu / Submenu / Command
- Menu item actions:
  - Copy content to clipboard
  - Ctrl + Left Click: Open path or URL
  - Ctrl + Right Click: Send content as keystrokes (SendKeys)

## How to Use

1. Place the compiled EXE in any folder.
2. Place one or more `.csv` files in the same folder or subfolders.
3. Each CSV file should follow this format:

Main Menu,Submenu,Display Name,Content


4. When launched, the tray icon will appear in the system tray.
5. Right-click the icon to reload the menu or exit the application.

## Notes

- CSV files must be UTF-8 encoded.
- The Content field must not be empty.
