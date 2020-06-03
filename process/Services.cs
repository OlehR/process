using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceProcess;
using System.Threading;

/// <summary>
/// Клас для управління Windows службою
/// </summary>
class Services
{
    private ServiceController sc;
    public string log;

    public void Log(string parLog = "")
    {
        DateTime varDT = DateTime.Now;
        log = varDT.ToString() + "=>" + (parLog.Length > 0 ? parLog + "=>" : "") + sc.Status.ToString() + "\n";
    }

    public Services(string varServiceName = "MSSQLServerOLAPService", string varServer = "localhost")
    {
        sc = new ServiceController(varServiceName, varServer);
    }

    public bool WaitServiceInProces(int parWaitSec = 200)
    {
        Log("WaitServiceInProces");
        do
        {
            Thread.Sleep(1000);
            sc.Refresh();
            Log();
        }
        while ((sc.Status == ServiceControllerStatus.StopPending || sc.Status == ServiceControllerStatus.StartPending) && parWaitSec-- > 0);
        return (parWaitSec > 0);
    }

    public bool Start(int parWaitSec = 200)
    {
        Log("Start");
        try
        {
            if (!WaitServiceInProces(parWaitSec)) return false;

            if (sc.Status == ServiceControllerStatus.Stopped)
            {
                sc.Start();
                do
                {
                    Thread.Sleep(1000);
                    sc.Refresh();
                    Log();
                }
                while (sc.Status != ServiceControllerStatus.Running && parWaitSec-- > 0);
            }
            return (parWaitSec > 0);
        }
        catch (Exception ex)
        {
            return false;
        }
    }

    public bool Stop(int parWaitSec = 200)
    {
        Log("Stop");
        try
        {
            if (!WaitServiceInProces(parWaitSec)) return false;

            if (sc.Status == ServiceControllerStatus.Running)
            {
                sc.Stop();
                do
                {
                    Thread.Sleep(1000);
                    sc.Refresh();
                    Log();
                }
                while (sc.Status != ServiceControllerStatus.Stopped && parWaitSec-- > 0);
            }
            return (parWaitSec > 0);
        }
        catch (Exception ex)
        {
            return false;
        }
    }
    public bool ReStart(int parWaitSec = 200)
    {
        log = "";
        return Stop(parWaitSec) && Start(parWaitSec);
    }
    public bool IsStart()
    {
        sc.Refresh();
        return (sc.Status == ServiceControllerStatus.Running);
    }

}