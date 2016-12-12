using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

namespace themechanger
{
    public partial class ThemeChanger : Form
    {
        readonly RegistryKey _rkApp = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
        private readonly List<string> _themes;

        bool _isRunStartup = false;

        public ThemeChanger()
        {
            InitializeComponent();

            notifyIcon1.ContextMenuStrip = contextMenuStrip1;

            this.changeThemeToolStripMenuItem.Click += changeToolStripMenuItem_Click;
            this.startupToolStripMenuItem.Click += startupToolStripMenuItem_Click;
            this.exitToolStripMenuItem.Click += exitToolStripMenuItem_Click;

            _isRunStartup = _rkApp.GetValue("ThemeChanger") != null;
            _themes = new ThemeFolder().GetThemes();

            ChangeStartupMenuText();
            ChangeTheme();
        }

        private bool _allowVisible;     // ContextMenu's Show command used
        private bool _allowClose;       // ContextMenu's Exit command used

        protected override void SetVisibleCore(bool value)
        {
            if (!_allowVisible)
            {
                value = false;
                if (!this.IsHandleCreated) CreateHandle();
            }
            base.SetVisibleCore(value);
        }

        private void ChangeStartupMenuText()
        {
            startupToolStripMenuItem.Text = _isRunStartup ? @"Remove Windows Startup" : @"Add Windows Startup";
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!_allowClose)
            {
                this.Hide();
                e.Cancel = true;
            }
            base.OnFormClosing(e);
        }

        private void ChangeTheme()
        {
            //method 1
            LaunchCommandLineApp("ThemeSwitcher.exe", "start \"");

            //method 2
            //new Theme().SwitchTheme(_themes[new Random().Next(0, _themes.Count)]);

            //method 3
            //LaunchCommandLineApp("ThemeTool.exe", "changetheme " + _themes[new Random().Next(0, _themes.Count)]);
        }

        private void startupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!_isRunStartup)
            {
                // Add the value in the registry so that the application runs at startup
                _rkApp.SetValue("ThemeChanger", Application.ExecutablePath.ToString());
                _isRunStartup = true;
            }
            else
            {
                // Remove the value from the registry so that the application doesn't start
                _rkApp.DeleteValue("ThemeChanger", false);
                _isRunStartup = false;
            }

            ChangeStartupMenuText();
        }

        private void changeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ChangeTheme();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _allowClose = true;
            Application.Exit();
        }

        private void LaunchCommandLineApp(string filename, string arguments)
        {
            // Use ProcessStartInfo class
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                CreateNoWindow = false,
                UseShellExecute = false,
                FileName = filename,
                WindowStyle = ProcessWindowStyle.Hidden,
                Arguments = arguments
            };

            try
            {
                // Start the process with the info we specified.
                // Call WaitForExit and then the using statement will close.
                using (Process exeProcess = Process.Start(startInfo))
                {
                    exeProcess.WaitForExit();
                }
            }
            catch
            {
                // Log error.
            }
        }
    }
}