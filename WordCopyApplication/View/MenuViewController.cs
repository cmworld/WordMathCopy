using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TYWordCopy.Controller;
using TYWordCopy.Model;
using TYWordCopy.Properties;
using TYWordCopy.Util;

namespace TYWordCopy.View
{
    class MenuViewController
    {
        private TYWordCopyAppController _controller;

        private ContextMenu contextMenu;
        private NotifyIcon _notifyIcon;

        public MenuViewController(TYWordCopyAppController controller)
        {
            this._controller = controller;

            LoadMenu();

            _notifyIcon = new NotifyIcon();
            _notifyIcon.Icon = Icon.FromHandle(Resources.icon16.GetHicon());

            Configuration config = new Configuration();
            string text = config.app_name + " " + Configuration.Version;

            _notifyIcon.Text = text.Substring(0, Math.Min(63, text.Length));
            _notifyIcon.Visible = true;
            _notifyIcon.ContextMenu = contextMenu;

            controller.Errored += controller_Errored;
        }

        private void LoadMenu()
        {
            this.contextMenu = new ContextMenu(new MenuItem[] {
                CreateMenuItem("About...", new EventHandler(this.CallBack_About)),
                new MenuItem("-"),
                CreateMenuItem("Quit", new EventHandler(this.CallBack_Quit))
            });
        }

        private void CallBack_About(object sender, EventArgs e)
        {
            Process.Start("http://ebangshou.me");
        }

        private void CallBack_Quit(object sender, EventArgs e)
        {
            _controller.Stop();
            _notifyIcon.Visible = false;
            Application.Exit();
        }

        private MenuItem CreateMenuItem(string text, EventHandler clickCallBack)
        {
            return new MenuItem(I18N.GetString(text), clickCallBack);
        }

        void controller_Errored(object sender, System.IO.ErrorEventArgs e)
        {
            MessageBox.Show(e.GetException().ToString(), String.Format(I18N.GetString("Error: {0}"), e.GetException().Message));
        }
    }
}
