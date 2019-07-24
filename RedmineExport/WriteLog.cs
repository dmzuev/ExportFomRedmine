using System;
using System.IO;

namespace SrvAsket
{
    class WriteLog
    {
        const string defaultLogFileName = "RedmineExport_#DateTimeCreare#.log";
        const string defaultRowLog = "#############################################################################";
        const string defaultRowSeparator = "----------------------------------------------------------------------------";
        static string logFileName = "";
        static StreamWriter sw;
        static Boolean writeLog = true;
        private WriteLog(){
 
        }

        public static WriteLog Start{
            get{
                return Nested.Start;
            }
        }
        public static void writeStringInLog(String logString)
        {
            string stringToWrite = formatRow(logString);            
            Console.WriteLine(stringToWrite);
            Console.WriteLine(defaultRowSeparator);
            if (writeLog){
                try
                {
                    sw = new StreamWriter(logFileName, true, System.Text.Encoding.Default);
                    {
                        sw.WriteLine(stringToWrite);
                        sw.WriteLine(defaultRowSeparator);
                        sw.Close();
                    }
                }
                catch (IOException ex)
                {
                    Console.WriteLine(formatRow("Ошибка записи события '" + logString + "'в лог файл"));
                    Console.WriteLine(defaultRowSeparator);
                    ex = null;

                }
            }
        }
        
        private static string formatRow(String logString)
        {
            return  DateTime.Now.ToString() + " : " + logString ;
        }
        private class Nested
        {
            static Nested()
            {
                string DateTimeCreateLogFile = DateTime.Now.ToString().Replace(':', '_').Replace('.', '_').Replace(' ', '_');
                logFileName = defaultLogFileName.Replace("#DateTimeCreare#", DateTimeCreateLogFile);
                Console.WriteLine(defaultRowLog);                    
                writeStringInLog( "Создан файл лога " + logFileName);                    
            }
            internal static readonly WriteLog Start = new WriteLog();
        }

    }
}

