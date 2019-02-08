using System;
using System.ServiceProcess;
using System.Diagnostics;
using System.Threading;

namespace ServiceControllerSample
{
    class Program
    {
        public enum SimpleServiceCustomCommands
        { StopWorker = 128, RestartWorker, CheckWorker };
        static void Main(string[] args)
        {
            ServiceController[] scServices;
            scServices = ServiceController.GetServices();

            foreach (ServiceController scTemp in scServices)
            {

                if (scTemp.ServiceName == "RemoteRegistry")
                {
                    // Display properties for the Simple Service sample
                    // from the ServiceBase example.
                    ServiceController sc = new ServiceController   ("RemoteRegistry" );
                    Console.WriteLine("Status = " + sc.Status);
                    Console.WriteLine("Can Pause and Continue = " + sc.CanPauseAndContinue);
                    Console.WriteLine("Can ShutDown = " + sc.CanShutdown);
                    Console.WriteLine("Can Stop = " + sc.CanStop);
                    if (sc.Status == ServiceControllerStatus.Stopped)
                    {
                        sc.Start();
                        while (sc.Status == ServiceControllerStatus.Stopped)
                        {
                            Thread.Sleep(1000);
                            sc.Refresh();
                        }
                    }
                    // Issue custom commands to the service
                    // enum SimpleServiceCustomCommands 
                    //    { StopWorker = 128, RestartWorker, CheckWorker };
                    sc.ExecuteCommand((int)SimpleServiceCustomCommands.StopWorker);
                    sc.ExecuteCommand((int)SimpleServiceCustomCommands.RestartWorker);
                    sc.Pause();
                    while (sc.Status != ServiceControllerStatus.Paused)
                    {
                        Thread.Sleep(1000);
                        sc.Refresh();
                    }
                    Console.WriteLine("Status = " + sc.Status);
                    sc.Continue();
                    while (sc.Status == ServiceControllerStatus.Paused)
                    {
                        Thread.Sleep(1000);
                        sc.Refresh();
                    }
                    Console.WriteLine("Status = " + sc.Status);
                    sc.Stop();
                    while (sc.Status != ServiceControllerStatus.Stopped)
                    {
                        Thread.Sleep(1000);
                        sc.Refresh();
                    }
                    Console.WriteLine("Status = " + sc.Status);
                    String[] argArray = new string[] { "ServiceController arg1", "ServiceController arg2" };
                    sc.Start(argArray);
                    while (sc.Status == ServiceControllerStatus.Stopped)
                    {
                        Thread.Sleep(1000);
                        sc.Refresh();
                    }
                    Console.WriteLine("Status = " + sc.Status);
                    // Display the event log entries for the custom commands
                    // and the start arguments.
                    EventLog el = new EventLog("Application");
                    EventLogEntryCollection elec = el.Entries;
                    foreach (EventLogEntry ele in elec)
                    {
                        if (ele.Source.IndexOf("SimpleService.OnCustomCommand") >= 0 |
                            ele.Source.IndexOf("SimpleService.Arguments") >= 0)
                            Console.WriteLine(ele.Message);
                    }
                }
            }


        }
    }
}