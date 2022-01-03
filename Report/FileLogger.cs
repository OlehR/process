using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace Utils
{
    public static class FileLogger
    {
        private static Dictionary<int, Type> _types = new Dictionary<int, Type>();

        private static readonly object Locker = new object();
        private static readonly object DictionaryLocker = new object();
        static StringBuilder Log = new StringBuilder();
        public static string GetLog{get{ return Log.ToString(); } }
        private static bool IsFirst = true;
        static void  CreateDir()
        {
            if (!Directory.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Log")))
                    Directory.CreateDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Log"));
        }
        public static void ExtLogForClass(Type type, int hashCode, string message, string parameters = null)
        {
            if (!string.IsNullOrWhiteSpace(parameters))
                message += $" {parameters}";

            WriteLogMessage($"[{type} - {hashCode}] {message}");
        }

        public static void ExtLogForClassConstruct(Type type, int hashCode, string parameters = null)
        {
            lock (DictionaryLocker)
            {
                if (!_types.ContainsKey(hashCode))
                    _types.Add(hashCode, type);
            }

            var message = "";
            if (!string.IsNullOrWhiteSpace(parameters))
                message += $" {parameters}";

            WriteLogMessage($"[{type} - {hashCode}] constructed {message}");
        }

        public static void ExtLogForClassDestruct(int hashCode, string parameters = null)
        {
            Type type;
            lock (DictionaryLocker)
            {
                if (_types.TryGetValue(hashCode, out type))
                    _types.Remove(hashCode);
            }

            var message = "";
            if (!string.IsNullOrWhiteSpace(parameters))
                message += $" {parameters}";

            WriteLogMessage($"[{type} - {hashCode}] destructed {message}");
        }

        public static void WriteLogMessage(this string message, string subLog = null)
        {
            
#if DEBUG
            message.WriteConsoleDebug();
#endif
            Task.Run(() =>
            {
                lock (Locker)
                {
                    if (IsFirst)
                    {
                        CreateDir();
                        IsFirst = false;
                    }
                    var date = DateTime.Now;
                    var str = $@"[{date:dd-MM-yyyy HH:mm:ss}] {message}{Environment.NewLine}";
                    Log.Append(str);
                    File.AppendAllText(
                        $"{Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"Log", $"{date.Year}{date.Month}{date.Day}.log")}",str);
                    Console.WriteLine(str);
                }
            });
        }

        public static void WriteConsoleDebug(this string message)
        {
            Console.WriteLine($@"[{DateTime.Now:dd-MM-yyyy HH:mm:ss}] {message}");
            // Console.ReadKey();
        }
    }
}