using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace DynamicOutlook
{
    /// <summary>
    /// Top-level object used to interop with Outlook
    /// </summary>
    class OutlookApplication : IDisposable
    {
        /// <summary>
        /// Gets an instance of Outlook that we can interact with
        /// </summary>
        private dynamic UnderlyingObject
        {
            get
            {
                if (_underlyingObject == null)
                {
                    try
                    {
                        // Try to get an already running instance
                        _underlyingObject = Marshal.GetActiveObject("Outlook.Application");
                    }
                    catch (COMException)
                    {
                        // Capture previously running instances of Outlook
                        Process[] previousOutlookProcesses = Process.GetProcessesByName("outlook");

                        // Wasn't running or permission denied, create our own
                        _underlyingObject = Activator.CreateInstance(OutlookApplicationType);

                        // Capture new list of outlook processes
                        Process[] currentOutlookProcesses = Process.GetProcessesByName("outlook");

                        // Find new Outlook process(es)
                        _ownedOutlookProcess = null;
                        foreach (Process currentOutlookProcess in currentOutlookProcesses)
                        {
                            if (!previousOutlookProcesses.Any(
                                previousOutlookProcess =>
                                {
                                    return (
                                        (previousOutlookProcess.Id == currentOutlookProcess.Id)
                                        && previousOutlookProcess.StartTime.Equals(currentOutlookProcess.StartTime));
                                }))
                            {
                                if (_ownedOutlookProcess == null)
                                {
                                    _ownedOutlookProcess = currentOutlookProcess;
                                }
                                else
                                {
                                    // We have more than one new Outlook process, so we can't tell which is which...
                                    _ownedOutlookProcess = null;
                                    break;
                                }
                            }
                        }
                    }
                }

                return _underlyingObject;
            }
        }
        private dynamic _underlyingObject = null;

        /// <summary>
        /// The Outlook process owned by this class, i.e. started and expected to be disposed of by this class
        /// </summary>
        private Process _ownedOutlookProcess = null;

        /// <summary>
        /// Gets the type associated with the Outlook ProgID (will return null if Outlook is not installed)
        /// </summary>
        public static Type OutlookApplicationType
        {
            get
            {
                return Type.GetTypeFromProgID("Outlook.Application");
            }
        }

        /// <summary>
        /// Ensures that we have a MAPI session active in Outlook
        /// </summary>
        private void EnsureMAPINamespace()
        {
            if (_MAPINamespace == null)
            {
                _MAPINamespace = UnderlyingObject.GetNamespace("MAPI");
            }
        }

        /// <summary>
        /// The MAPI session
        /// </summary>
        private dynamic _MAPINamespace = null;

        /// <summary>
        /// Creates a new mail item
        /// </summary>
        /// <returns></returns>
        public OutlookMailItem CreateMailItem()
        {
            EnsureMAPINamespace();

            // olMailItem == 0
            return new OutlookMailItem(UnderlyingObject.CreateItem(0));
        }

        /// <summary>
        /// Invokes the Application.Quit method. Note that Application.Quit doesnt work reliably in Outlook, so we're making this async
        /// and if it fails to exit in a reasonable amount of time (including process going away) we instead leave Outlook in the showing
        /// state, so it's at least obvious to the user that Outlook is still running.
        /// </summary>
        private async Task QuitAsync()
        {
            // Ask Outlook to quit nicely
            UnderlyingObject.Quit();

            // Wait for some time for the process to actually exit
            await Task.Run(
                () =>
                {
                    if (!_ownedOutlookProcess.WaitForExit(5000))
                    {
                        // Outlook hasn't quit. At this point we can either kill the process or show the main window.
                        // Since Outlook is prone to corrupting its database on forceful exit, we opt for the lessoptimal latter option.
                        EnsureMAPINamespace();
                        _MAPINamespace.GetDefaultFolder(6).Display(); // 6 == olInboxFolder
                        ReleaseMAPINamespace();
                    }
                });

            // Make sure to clean up references
            ReleaseUnderlyingObject();
        }

        /// <summary>
        /// Releases the login session
        /// </summary>
        private void ReleaseMAPINamespace()
        {
            if (_MAPINamespace != null)
            {
                Marshal.ReleaseComObject(_MAPINamespace);
                _MAPINamespace = null;
            }
        }

        /// <summary>
        /// Releases the Outlook application
        /// </summary>
        private void ReleaseUnderlyingObject()
        {
            if (_underlyingObject != null)
            {
                Marshal.ReleaseComObject(_underlyingObject);
                _underlyingObject = null;

            }
        }

        public void Dispose()
        {
            if (_underlyingObject != null)
            {
                ReleaseMAPINamespace();

                if (_ownedOutlookProcess != null)
                {
                    QuitAsync();
                }
                else
                {
                    ReleaseUnderlyingObject();
                }
            }
        }
    }
}
