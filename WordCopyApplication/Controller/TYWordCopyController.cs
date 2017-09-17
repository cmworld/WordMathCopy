using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using TYWordCopy.Controller.Server;
using TYWordCopy.Model;

namespace TYWordCopy.Controller
{
    class TYWordCopyAppController
    {
        private Thread _ramThread;
        private bool stopped = false;

        private DataConvert dataConvert;
        private FocusMontorer focusMontorer;
        private ClipboardMonitor clipboardMontorer;

        public event ErrorEventHandler Errored;
        public event EventHandler<ClipboardMonitor.ClipboardChangedEventArgs> ClipboardChanged;
        public event EventHandler<DataConvert.DataConvertEventArgs> DataConverted;
        public event EventHandler<System.Windows.Automation.AutomationFocusChangedEventArgs> FocusChanged;

        public TYWordCopyAppController()
        {

            Logging.OpenLogFile();

            //TODO: load config

            StartReleasingMemory();
        }

        protected void ReportError(Exception e)
        {
            if (Errored != null)
            {
                Errored(this, new ErrorEventArgs(e));
            }
        }

        public void Start()
        {
            _Reload();
        }

        public void Stop()
        {
            if (stopped)
            {
                return;
            }
            stopped = true;
        }

        private void _Reload()
        {

            try
            {
                if (dataConvert == null)
                {
                    dataConvert = new DataConvert(this);
                    dataConvert.DataConverted += data_converted;
                }

                if (focusMontorer == null)
                {
                    focusMontorer = new FocusMontorer();
                    focusMontorer.FocusChanged += focusMontorer_FocusChanged;
                }

                if (clipboardMontorer == null)
                {
                    clipboardMontorer = new ClipboardMonitor(this);
                    clipboardMontorer.ClipboardChanged += clipboardMontorer_changed;
                }

            }
            catch( Exception e) {

                Logging.LogUsefulException(e);
                ReportError(e);
            }
        }

        private void clipboardMontorer_changed(object sender, ClipboardMonitor.ClipboardChangedEventArgs e)
        {
            if (ClipboardChanged != null)
                ClipboardChanged(this, e);
        }

        private void data_converted(object sender, DataConvert.DataConvertEventArgs e)
        {
            if (DataConverted != null)
                DataConverted(this, e);
        }

        private void focusMontorer_FocusChanged(object sender, System.Windows.Automation.AutomationFocusChangedEventArgs e)
        {
            if (FocusChanged != null)
                FocusChanged(this, e);       
        }

        private void StartReleasingMemory()
        {
            _ramThread = new Thread(new ThreadStart(ReleaseMemory));
            _ramThread.IsBackground = true;
            _ramThread.Start();
        }

        private void ReleaseMemory()
        {
            while (true)
            {
                Util.Utils.ReleaseMemory(false);
                Thread.Sleep(30 * 1000);
            }
        }
    }
}
