using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Messaging;


namespace sigConfigClientService
{

    // Variables used by the Windows service code
    public enum ServiceState
    {
        SERVICE_STOPPED = 0X00000001,
        SERVICE_START_PENDING = 0x00000002,
        SERVICE_STOP_PENDING = 0x00000003,
        SERVICE_RUNNING = 0x00000004,
        SERVICE_CONTINUE_PENDING = 0x00000005,
        SERVICE_PAUSE_PENDING = 0x00000006,
        SERVICE_PAUSED = 0x00000007,
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct ServiceStatus
    {
        public long dwServiceType;
        public ServiceState dwCurrentState;
        public long dwControlsAccepted;
        public long dwWin32ExitCode;
        public long dwServiceSpecificExitCode;
        public long dwCheckPoint;
        public long dwWaitHint;
    };

    public partial class SigConfigClient : ServiceBase
    {

        // Importing function for setting service status. 
        [DllImport("advapi32.dll", SetLastError=true)]
            private static extern bool SetServiceStatus(IntPtr handle, ref ServiceStatus serviceStatus);

        public SigConfigClient(string[] args)
        {

            // Don't log automatically!
            this.AutoLog = false;

            // Initialise variables and check commandline arguments for values (for manual specific logging)
            string eventSourceName = "sigConfigClientSource";
            string logName = "sigConfigClientLog";
            if (args.Count() > 0)
            {
                eventSourceName = args[0];
            }
            if (args.Count() > 1)
            {
                logName = args[1];
            }

            // Set up the logging
            InitializeComponent();
            sigConfigClientServiceLog = new System.Diagnostics.EventLog();
            if (!System.Diagnostics.EventLog.SourceExists(eventSourceName))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    eventSourceName, logName);
            }
            sigConfigClientServiceLog.Source = eventSourceName;
            sigConfigClientServiceLog.Log = logName;
            
            // Logging
            sigConfigClientServiceLog.WriteEntry("Roswell Email Signature Sync service (client mode) created.");

            // TODO: CREATE MQ FOR SERVICE HERE
        }

        protected override void OnStart(string[] args)
        {
            sigConfigClientServiceLog.WriteEntry("In OnStart");

            // Setting up a timer to trigger every minute.
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 60000;
            timer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimer);
            timer.Start();

            // Update the service state to START_PENDING
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            // When the service runs, set the state to 'running'.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            // TODO: Check here that the MQ is still alive. If it isn't, wake it up. 

            // TODO: Check any messages in the queue

            // TODO: For any messages we find, process them with [FUNCTION TO BE WRITTEN]().
        }

        protected override void OnStop()
        {
            sigConfigClientServiceLog.WriteEntry("Stopping.");

            // Update the service state to pend ending.
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_STOP_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            // Update the service state to Stopped.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_STOPPED;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
        }

        public void OnTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            // TODO: Insert monitoring activities here.
            sigConfigClientServiceLog.WriteEntry("Service woken by timer, performing sync and monitoring if necessary.", EventLogEntryType.Information);

            // TODO: check file status when timer activates.
            /* PSEUDOCODE:
             * 
             * get current status
             * if not running:
             *     Start service
             * check file status
             * if altered since last activation:
             *      perform 365 update
             *      send messages to sigConfigClientServices to perform local updates.
             * if old status not running:
             *      set status to previously running status
             * 
             */

        }

        // TODO: When do we use this?
        protected override void OnContinue()
        {
            sigConfigClientServiceLog.WriteEntry("Continuing.");
        }

    }
}
