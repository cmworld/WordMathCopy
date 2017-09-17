using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using TYWordCopy.Controller;
using TYWordCopy.Util;
using TYWordCopy.View;

namespace TYWordCopy
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {

            Util.Utils.ReleaseMemory(true);
            using (Mutex mutex = new Mutex(false, "Global\\TYWordCopyApplication_" + Application.StartupPath.GetHashCode()))
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                if (!mutex.WaitOne(0, false))
                {
                    Process[] oldProcesses = Process.GetProcessesByName("TYWordCopyApplication");
                    if (oldProcesses.Length > 0)
                    {
                        Process oldProcess = oldProcesses[0];
                    }
                    MessageBox.Show(I18N.GetString("程序以运行"));
                    return;
                }
                Directory.SetCurrentDirectory(Application.StartupPath);

                TYWordCopyAppController controller = new TYWordCopyAppController();

                MenuViewController viewController = new MenuViewController(controller);

                controller.Start();

                Application.Run();
            }
        }
    }
}
