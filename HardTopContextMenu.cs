#region Using statements

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

#endregion Using statements

namespace HardTop
{
    internal class HardTopContextMenu : IDisposable
    {
        #region Internal context menu

        internal ContextMenu ContextMenu { get; private set; }
        internal const int NUMBER_OF_FIXED_ITEMS = 9;

        #endregion Internal context menu

        #region Internal contructor

        internal HardTopContextMenu()
        {
            CreateContextMenu();
            ContextMenu.Collapse += ContextMenu_Collapse;
            ContextMenu.Popup += ContextMenu_Popup;
            AddUpdatedWindowsInfoToContextMenu();
        }

        #endregion Internal constructor

        #region Event handling

        private void ContextMenu_Popup(object sender, EventArgs e)
        {
            AddUpdatedWindowsInfoToContextMenu();
        }

        private void ContextMenu_Collapse(object sender, EventArgs e)
        {
            AddUpdatedWindowsInfoToContextMenu();
        }

        private void WindowItem_Click(object o, EventArgs e)
        {
            MenuItem mi = null;
            try
            {
                mi = (MenuItem)o;
                mi.Enabled = false;
                mi.Checked = !mi.Checked;
                NativeMethods.ToggleWindowAlwaysOnTop((IntPtr)mi.Tag, mi.Checked);
            }
            catch (Exception ex)
            {
                Message.Show(Resources.UnhandledException, ex);
            }
            finally
            {
                EnableMenuItem(mi);
            }
        }

        private static void ExitMenuClick(object o, EventArgs e)
        {
            Application.Exit();
        }

        private void StartWithWindowsClick(object o, EventArgs e)
        {
            try
            {
                var mi = (MenuItem)o;
                mi.Enabled = false;
                mi.Checked = !mi.Checked;
                Settings.StartWithWindows = mi.Checked;
                EnableMenuItem(mi);
            }
            catch (Exception ex)
            {
                Message.Show(Resources.FailedToUpdateRegistry, ex);
            }
        }

        private void AboutClick(object o, EventArgs e)
        {
            var mi = (MenuItem)o;
            mi.Enabled = false;
            Message.Show(Resources.About);
            EnableMenuItem(mi);
        }

        #endregion Event handling

        #region Private method for context menu creation

        private void CreateContextMenu() => ContextMenu = new ContextMenu(new MenuItem[NUMBER_OF_FIXED_ITEMS]
        {
            new MenuItem(Resources.AboutMenu, AboutClick),
            new MenuItem(Resources.SeparatorMenu),
            new MenuItem(Resources.DonationMenu, (o, e) => { new Process() { StartInfo = new ProcessStartInfo(Resources.DonationUrl) { UseShellExecute = true } }.Start(); } ),
            new MenuItem(Resources.SeparatorMenu),
            new MenuItem(Resources.StartWithWindowsMenu, StartWithWindowsClick) { Checked = Settings.StartWithWindows },
            new MenuItem(Resources.SeparatorMenu),
            new MenuItem(Resources.ExitMenu, ExitMenuClick),
            new MenuItem(Resources.WindowsMenu) { DefaultItem = true, BarBreak = true },
            new MenuItem(Resources.SeparatorMenu)
        });

        #endregion Private method for context menu creation

        #region Private helper methods for menu items

        private void AddUpdatedWindowsInfoToContextMenu()
        {
            try
            {
                while (ContextMenu.MenuItems.Count > NUMBER_OF_FIXED_ITEMS) ContextMenu.MenuItems.RemoveAt(NUMBER_OF_FIXED_ITEMS);
                NativeMethods.GetDesktopWindowHandlesAndTitles(out List<IntPtr> handles, out List<string> titles);
                List<string> ignoreTheseWindows = new List<string>() { "Program Manager", "MainWindow", "Snipping Tool", "Task Manager", "Task Manager Properties" };
                for (int i = 0; i < titles.Count; i++)
                    if (!ignoreTheseWindows.Contains(titles[i]))
                        ContextMenu.MenuItems.Add(new MenuItem(titles[i], WindowItem_Click) { Name = titles[i], Tag = handles?[i], Checked = NativeMethods.AlwaysOnTopWindows().Contains((IntPtr)handles?[i]) });
            }
            catch (Exception ex)
            {
                Message.Show(Resources.UnhandledException, ex);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static void EnableMenuItem(MenuItem mi)
        {
            if (mi != null) mi.Enabled = true;
        }

        #endregion Private helper methods for menu items

        #region IDisposable methods

        /// <summary>
        /// Disposes the context menu
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposing)
            {
                return;
            }
            CleanupContextMenu();
        }

        private void CleanupContextMenu()
        {
            ContextMenu?.Dispose();
            ContextMenu = null;
        }

        #endregion IDisposable methods
    }
}