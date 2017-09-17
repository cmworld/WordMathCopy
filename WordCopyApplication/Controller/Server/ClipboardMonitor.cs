using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TYWordCopy.Util;

namespace TYWordCopy.Controller.Server
{

    class ClipboardMonitor : ContainerControl
    {

        private IntPtr nextClipboardViewer;
        private TYWordCopyAppController _controller;
        public event EventHandler<ClipboardChangedEventArgs> ClipboardChanged;
        IDataObject iData;

        public class ClipboardChangedEventArgs : EventArgs
        {
            public readonly IDataObject DataObject;

            public ClipboardChangedEventArgs(IDataObject dataObject)
            {
                DataObject = dataObject;
            }
        }

        public ClipboardMonitor(TYWordCopyAppController controller)
        {
            TYWordCopyAppController _controller = controller;
            controller.DataConverted += resetClipboardData;

            nextClipboardViewer = (IntPtr)SetClipboardViewer((int)this.Handle);
        }

        [DllImport("user32.dll")]
        protected static extern int SetClipboardViewer(int hWndNewViewer);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern bool ChangeClipboardChain(IntPtr hWndRemove, IntPtr hWndNewNext);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SendMessage(IntPtr hwnd, int wMsg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern IntPtr GetClipboardOwner();

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("kernel32.dll")]
        static extern uint GetCurrentProcessId();

        protected override void WndProc(ref Message m)
        {
            // defined in winuser.h
            const int WM_DRAWCLIPBOARD = 0x308;
            const int WM_CHANGECBCHAIN = 0x030D;
            const int WM_PASTE = 0x0302;

            switch (m.Msg)
            {
                case WM_DRAWCLIPBOARD:
                    uint pid; uint cpid;
                    GetWindowThreadProcessId(GetClipboardOwner(), out pid);
                    cpid = GetCurrentProcessId();

                    // i only want info about what is copied by other programs.
                    if (pid != cpid)
                    {
                        OnClipboardChanged();
                    }

                    SendMessage(nextClipboardViewer, m.Msg, m.WParam, m.LParam);
                    break;

                case WM_CHANGECBCHAIN:
                    if (m.WParam == nextClipboardViewer)
                        nextClipboardViewer = m.LParam;
                    else
                        SendMessage(nextClipboardViewer, m.Msg, m.WParam, m.LParam);
                    break;
                case WM_PASTE:
                    break;

                default:
                    base.WndProc(ref m);
                    break;
            }
        }

        void OnClipboardChanged()
        {
            iData = Clipboard.GetDataObject();
            if (ClipboardChanged != null)
            {
                ClipboardChanged(this, new ClipboardChangedEventArgs(iData));
            }
        }

        void resetClipboardData(object sender, DataConvert.DataConvertEventArgs e)
        {
            Clipboard.Clear();

            var dataObject = new DataObject();
            dataObject.SetData(DataFormats.Rtf, e.Rtf);

            var htmlFragment = ClipboardHelper.GetHtmlDataString(e.Html);

            // re-encode the string so it will work  correctly (fixed in CLR 4.0)      
            if (Environment.Version.Major < 4 && e.Html.Length != Encoding.UTF8.GetByteCount(e.Html))
                htmlFragment = Encoding.Default.GetString(Encoding.UTF8.GetBytes(htmlFragment));

            dataObject.SetData(DataFormats.Html, htmlFragment);

            using (RichTextBox rtBox = new RichTextBox())
            {
                rtBox.Rtf = e.Rtf;
                string plainText = rtBox.Text;

                plainText = plainText.Replace("\r\n", " ");

                dataObject.SetData(DataFormats.Text, plainText);
                dataObject.SetData(DataFormats.UnicodeText, plainText);
            }

            Clipboard.SetDataObject(dataObject, true);
        }
    }
}
