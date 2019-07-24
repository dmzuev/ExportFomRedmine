using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace SrvAsket

{
    class Program
    {
        public static bool ExitKey = false;
        static void Main(string[] args)
        {

            WriteLog LogStar = WriteLog.Start;
            Thread Redmine = new Thread(RunQueryToBitrix);
            Redmine.Start();
            WriteLog.writeStringInLog("Начало экспорта данных из Redmine");            
            while (!(ExitKey))
            {
                
                //ConsoleKeyInfo Key = Console.ReadKey(true);
                //if ((Key.Key == ConsoleKey.Q) && (Key.Modifiers == ConsoleModifiers.Control))
                //{
                //    ExitKey = true;
                //}
                if (!(Redmine.IsAlive))
                {
                    ExitKey = true;
                }
            }
            WriteLog.writeStringInLog("Плановое завершение программы");
        }
        private static void RunQueryToBitrix()
        {
            Redmine redmine = new Redmine();            
        }

    }
}
