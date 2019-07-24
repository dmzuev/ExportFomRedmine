using System;
using System.Data.SqlClient;

namespace SrvAsket { 

    public sealed class DB
    {
        public const string TABLE_REPORT_DEVELOPER = " stMain.dbo.АскетПорученияРазработка ";
        public const string TABLE_REDMINE_TIME_ENTRIES = " stMain.dbo.АскетСинхронизацияRedmine_Трудозатраты ";
        public const string TABLE_REDMINE_ISSUES = " stMain.dbo.АскетСинхронизацияRedmine_Задачи ";
        public const string TABLE_REDMINE_USERS = " stMain.dbo.АскетСинхронизацияRedmine_Пользователи ";
        public const string TABLE_ASKET_USERS = " stMain.dbo.Фамилии";

        private static volatile DB instance;
        private static object syncRoot = new Object();
        private static string connectionString = "Provider=SQLOLEDB;server=srv-asket;Initial Catalog=stMain;Persist Security Info=True;Integrated Security=SSPI";
        private static ADODB.Connection CnMain = new ADODB.Connection();
        private static SqlConnection connection = new SqlConnection("Data source=sql-analytics\\asket;Initial Catalog=stMain;Integrated Security=SSPI");
        private DB() { }

        public static DB Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (syncRoot)
                    {
                        if (instance == null)
                        {
                            if (Connect()==true)
                            {
                                instance = new DB();
                            }
                        }
                    }
                }
                return instance;
            }
        }
        private static bool Connect()
        {
            connection.Open();
            CnMain.ConnectionString = connectionString;
            try {
                CnMain.Open();
                if (TestQuery())
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch   
            {                
                WriteLog.writeStringInLog("Не смогли подключиться к БД. Строка подключения " + connectionString);
                return false;
            }
            
        }
        private static bool TestQuery()
        {
            if (CnMain.State == 0)
            {
                return false;
            }
            else
            {
                if (ExecuteQuery("Select top 1 Имя From stmain.dbo.Параметры"))
                {                   
                    WriteLog.writeStringInLog("Выполнили тестовый запрос");
                    return true;
                }
                else
                {                    
                    WriteLog.writeStringInLog("Тестовый запрос не выполнен");
                    return false;
                }
            }
        }
        //TODO: Дописать экранирование спецсимволов
        public static bool ExecuteQuery(string sql)
        {
            ADODB.Command cmd = new ADODB.Command
            {
                ActiveConnection = CnMain,
                CommandText = sql,
                CommandType = ADODB.CommandTypeEnum.adCmdText
            };
            try
            {
                object nRecordsAffected = Type.Missing;
                object oParams = Type.Missing;
                cmd.Execute(out nRecordsAffected, ref oParams, (int)ADODB.ExecuteOptionEnum.adExecuteNoRecords);
                return true;
            }
            catch
            {                
                WriteLog.writeStringInLog("Ошибка выполнения запроса: " + sql);
                return false;
            }        
        }
        //TODO: Дописать экранирование спецсимволов
        public static ADODB.Recordset Select(string sql)
        {
            ADODB.Recordset rs = new ADODB.Recordset
            {
                CursorLocation = ADODB.CursorLocationEnum.adUseClient
            };
            try
            { 
                rs.Open(sql, CnMain, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly);
                return rs;
            }
            catch
            {                
                WriteLog.writeStringInLog("Ошибка выполнения запроса: " + sql);
                return null;
            }
        
        }
        public static ADODB.Recordset Update(string sql)
        {
            ADODB.Recordset rs = new ADODB.Recordset { 
                CursorLocation = ADODB.CursorLocationEnum.adUseClient
            };
            try
            {
                rs.Open(sql, CnMain, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic);
                return rs;
            }
            catch
            {                
                WriteLog.writeStringInLog("Ошибка выполнения запроса: " + sql);
                return null;
            }
        
        }
    }
}