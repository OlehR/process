using System.Runtime.InteropServices;
[Guid("B992D892-36FE-46da-9B74-D3FE3567B37F")]
public interface IMyCOM_PROCESS
{
    [DispId(1)]
    void Process(string parServer, string parDB, string parCube,int Metod=0);
}

