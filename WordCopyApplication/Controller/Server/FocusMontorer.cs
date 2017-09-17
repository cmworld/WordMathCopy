using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Windows.Automation;

namespace TYWordCopy.Controller.Server
{

    class FocusMontorer
    {

        string[] browserNames;

        public event EventHandler<AutomationFocusChangedEventArgs> FocusChanged;
        AutomationFocusChangedEventHandler focusHandler;

        public FocusMontorer()
        {
            RegistryKey browserKeys;

            //on 64bit the browsers are in a different location
            browserKeys = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\WOW6432Node\Clients\StartMenuInternet");
            if (browserKeys == null)
                browserKeys = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Clients\StartMenuInternet");

            browserNames = browserKeys.GetSubKeyNames();
            
            focusHandler = new AutomationFocusChangedEventHandler(OnFocusChangedHandler);
            Automation.AddAutomationFocusChangedEventHandler(focusHandler);

        }

        private bool isBrowserFocus(string name)
        {
            if (browserNames.Length > 0)
            {
                foreach (string str in browserNames)
                {

                    if (str.ToLower().IndexOf(name) > -1)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private void OnFocusChangedHandler(object sender, AutomationFocusChangedEventArgs e)
        {
            AutomationElement element = sender as AutomationElement;
            if (element != null)
            {
                try
                {
                    //string name = element.Current.Name;
                    //string id = element.Current.AutomationId;
                    int processId = element.Current.ProcessId;
                    using (Process process = Process.GetProcessById(processId))
                    {
                        if (isBrowserFocus(process.ProcessName))
                        { 
                            if (FocusChanged != null)
                            {
                                FocusChanged(sender, e);
                            }
                        }

                        //Console.WriteLine("  Name: {0}, Id: {1}, Process: {2}", name, id, process.ProcessName);
                    }
                }
                catch //ElementNotAvailableException
                {

                }
            }
        }
    }
}
