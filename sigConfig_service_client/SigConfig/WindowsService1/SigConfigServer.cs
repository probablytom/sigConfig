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

namespace WindowsService1
{

    public enum ServiceState
    {
        SERVICE_STOPPED = 0x00000001,
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

    public partial class SigConfigServer : ServiceBase
    {

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(IntPtr handle, ref ServiceStatus serviceStatus);

        public SigConfigServer(String[] args)
        {
            InitializeComponent();
            this.AutoLog = false;
            sigConfigServerServiceLog = new System.Diagnostics.EventLog();
            string logSource = "sigConfigServerSource";
            string logName = "sigConfigServerLog";

            if (args.Count() > 0)
            {
                logSource = args[0];
            }
            if (args.Count() > 1)
            {
                logSource = args[1];
            }

            if (!System.Diagnostics.EventLog.SourceExists(logSource))
            {
                System.Diagnostics.EventLog.CreateEventSource(logSource, logName);
            }

            sigConfigServerServiceLog.Source = logSource;
            sigConfigServerServiceLog.Log = logName;
        }

        protected override void OnStart(string[] args)
        {

            // Update sigConfigServerService status to Start Pending.
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            // Update sigconfigServerService status to running. 
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            // Log that the service has begun. 
            sigConfigServerServiceLog.WriteEntry("Started sigConfig server.");
   
            // TODO: Check file status on service start and perform the first iteration of updates if so. 
        }

        protected override void OnStop()
        {

            // Set service status to pending stop
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_STOP_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            // Set service status to stopped.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_STOPPED;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            sigConfigServerServiceLog.WriteEntry("Stopped sigConfig server.");
            
        }

        protected override void OnContinue()
        {
            base.OnContinue();
            sigConfigServerServiceLog.WriteEntry("In OnContinue()");
        }

        protected void onTimer()
        {
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
            sigConfigServerServiceLog.WriteEntry("Timer activated, no checking implemented!", EventLogEntryType.Information);
        }

    }

}
