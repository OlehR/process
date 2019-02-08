using System.Runtime.InteropServices;
/// <summary>
/// Мой класс
/// </summary>
[Guid("107E60CC-08A3-4e0e-833E-D766ABDC7A6A"),
ClassInterface(ClassInterfaceType.None),
ComSourceInterfaces(typeof(IMyEvents))]
public class COM_PROCESS : IMyCOM_PROCESS
{
    /// <summary>
    /// Конструктор
    /// </summary>
    public COM_PROCESS()
    {
 
    }
    /// <summary>
    /// Привет!
    /// </summary>
    public void Process(string parServer, string parDB, string parCube,int Metod=0)
    {
//        MessageBox.Show((mymessage.Equals(String.Empty) ? "Привет!" : "Привет " + parServer), "Тест библиотека", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}


