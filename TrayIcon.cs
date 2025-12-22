using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic.FileIO;
using static NativeMethods;

namespace TrayIcon
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new TrayIconApplicationContext());
        }
    }

    public class TrayIconApplicationContext : ApplicationContext
    {
        private NotifyIcon _trayIcon;
        private ContextMenuStrip _trayMenu;

        public TrayIconApplicationContext()
        {
            _trayMenu = new ContextMenuStrip();
            ReloadTrayMenu();

            _trayIcon = new NotifyIcon
            {
                ContextMenuStrip = _trayMenu,
                Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath),
                Text = "Enjoy your day!",
                Visible = true
            };
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_trayIcon != null)
                {
                    _trayIcon.Visible = false;
                    _trayIcon.Dispose();
                }
                _trayMenu?.Dispose();
            }
            base.Dispose(disposing);
        }

        private void ApplicationExit()
        {
            Application.Exit();
        }

        private async void OnCommandMenuClick(object sender, MouseEventArgs e, TrayMenuItem item)
        {
            string errorMessage = "";
            try
            {
                if (!string.IsNullOrEmpty(item.ContentItem))
                {
                    errorMessage = "Failed to copy to clipboard: ";
                    Clipboard.SetText(item.ContentItem);

                    if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
                    {
                        if (e.Button == MouseButtons.Left)
                        {
                            errorMessage = "Failed to open: ";
                            Process.Start(new ProcessStartInfo(item.ContentItem) { UseShellExecute = true });
                        }
                        else if (e.Button == MouseButtons.Right)
                        {
                            errorMessage = "Failed to sendkey: ";
                            IntPtr hWnd = FindTopmostValidWindow();
                            if (hWnd != IntPtr.Zero)
                            {
                                SetForegroundWindow(hWnd);
                                await Task.Delay(1000);
                                SendKeys.SendWait(item.ContentItem);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(errorMessage + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static IntPtr FindTopmostValidWindow()
        {
            IntPtr result = IntPtr.Zero;
            int currentPid = Process.GetCurrentProcess().Id;

            EnumWindows((hWnd, lParam) =>
            {
                if (!IsWindowVisible(hWnd)) return true;

                StringBuilder className = new StringBuilder(256);
                GetClassName(hWnd, className, className.Capacity);
                if (className.ToString() == "Shell_TrayWnd") return true;

                GetWindowThreadProcessId(hWnd, out int pid);
                if (pid == currentPid) return true;

                StringBuilder title = new StringBuilder(256);
                GetWindowText(hWnd, title, title.Capacity);
                if (string.IsNullOrWhiteSpace(title.ToString())) return true;

                result = hWnd;
                return false;
            }, IntPtr.Zero);

            return result;
        }

        private void ReloadTrayMenu()
        {
            _trayMenu.Items.Clear();
            var loadItems = LoadFolder();

            var groupedItems = loadItems
                .GroupBy(x => x.MainMenu)
                .ToDictionary(
                    g => g.Key,
                    g => g.GroupBy(x => x.SubMenu)
                          .ToDictionary(sg => sg.Key, sg => sg.ToList())
                );

            foreach (var mainMenu in groupedItems)
            {
                ToolStripMenuItem mainMenuItem = null;

                foreach (var subMenu in mainMenu.Value)
                {
                    ToolStripMenuItem subMenuItem = null;

                    foreach (var item in subMenu.Value)
                    {
                        var displayItem = new ToolStripMenuItem(item.DisplayItem);
                        displayItem.MouseUp += (s, e) => OnCommandMenuClick(s, e, item);

                        subMenuItem ??= new ToolStripMenuItem(subMenu.Key);
                        subMenuItem.DropDownItems.Add(displayItem);
                    }

                    mainMenuItem ??= new ToolStripMenuItem(mainMenu.Key);
                    mainMenuItem.DropDownItems.Add(subMenuItem);
                }

                _trayMenu.Items.Add(mainMenuItem);
            }

            if (_trayMenu.Items.Count > 0)
            {
                _trayMenu.Items.Add(new ToolStripSeparator());
            }

            var reloadItem = new ToolStripMenuItem("Reload");
            reloadItem.Click += (s, e) => ReloadTrayMenu();
            _trayMenu.Items.Add(reloadItem);

            var exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += (s, e) => ApplicationExit();
            _trayMenu.Items.Add(exitItem);
        }

        private List<TrayMenuItem> LoadFolder()
        {
            var items = new List<TrayMenuItem>();

            foreach (var file in Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "*.csv", SearchOption.AllDirectories))
            {
                items.AddRange(LoadFile(file));
            }

            return items
                .OrderBy(x => x.MainMenu)
                .ThenBy(x => x.SubMenu)
                .ThenBy(x => x.DisplayItem)
                .ToList();
        }

        private IEnumerable<TrayMenuItem> LoadFile(string filePath)
        {
            var items = new List<TrayMenuItem>();

            try
            {
                using var csvFile = new TextFieldParser(filePath)
                {
                    HasFieldsEnclosedInQuotes = true,
                    TextFieldType = FieldType.Delimited,
                    TrimWhiteSpace = true
                };
                csvFile.SetDelimiters(",");

                while (!csvFile.EndOfData)
                {
                    var fields = csvFile.ReadFields();
                    if (fields.Length == 4)
                    {
                        var item = new TrayMenuItem
                        {
                            MainMenu = fields[0].Trim(),
                            SubMenu = fields[1].Trim(),
                            DisplayItem = fields[2].Trim(),
                            ContentItem = fields[3].Trim()
                        };

                        if (!string.IsNullOrEmpty(item.MainMenu) &&
                            !string.IsNullOrEmpty(item.DisplayItem) &&
                            !string.IsNullOrEmpty(item.ContentItem))
                        {
                            items.Add(item);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load file '{filePath}': {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ApplicationExit();
            }

            return items;
        }

        private class TrayMenuItem
        {
            public string MainMenu { get; set; }
            public string SubMenu { get; set; }
            public string DisplayItem { get; set; }
            public string ContentItem { get; set; }
        }
    }
}
