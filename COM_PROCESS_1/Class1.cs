using System;
using System.Runtime.InteropServices;
using Microsoft.AnalysisServices;

namespace COM_PROCESS
{
    [Guid("19B5E7C3-F37C-4cd1-A32F-76B2696EB473")]
	[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface _COM_PROCESS
	{
		[DispId(1)]
		int Process(string parServer, string parDB, string parCube,int parMetod);
	}

    [Guid("56340A88-8DA9-47d8-BD2F-1226E683A5B1")]
	[ClassInterface(ClassInterfaceType.None)]
    [ProgId("COM_PROCESS.PROCESS")]
	public class COM_PROCESS : _COM_PROCESS
	{
        public COM_PROCESS() { }
		public int Process(string parServer, string parDB, string parCube,int parMetod)
		{
            Server s = new Server();
            try
            {
                s.Connect(@"Data Source=" + parServer + ";Provider=msolap;");
                Database varDB = s.Databases.FindByName(parDB);
                if (parMetod == 1)
                {

                }
                varDB.Cubes.FindByName(parCube).Process(ProcessType.ProcessFull);
                return 1;
            }
            catch
            {
                return 0;
            }
		}

		
	}
}