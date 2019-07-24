using System;
using System.Xml;
using System.Collections;
using System.Collections.Generic;
using System.Net;
using System.IO;
using System.Text;
using System.Data;


namespace SrvAsket
{
    class RedmineEntity
    {
        const string REDMINE_URL = "http://redmine/";
        const string REDMINE_API_KEY = "&key=fe12556e4e0ed5c404c83fcf5ee44ddde35d8e4c";
        public const int Limit = 100;
        public static string Table { get; set; }
        public static string Field { get; set; }

        public static XmlReader GetXMLFromRedmine(string Query)
        {
            XmlReader GetXMLFromRedmine = XmlReader.Create(REDMINE_URL + Query + REDMINE_API_KEY);
            return GetXMLFromRedmine;
        }
        public virtual bool ReadPageFromRedmine(string Page)
        {

            XmlDocument xmlDocument = new XmlDocument();
            bool Result = false;
            string query = Page + "limit=" + Limit.ToString();
            xmlDocument.Load(GetXMLFromRedmine(query));
            int CountPage = (Convert.ToInt32(xmlDocument.DocumentElement.Attributes.GetNamedItem("total_count").InnerText) / Limit) + 1;
            int CountSavePage = 0;
            for (int i = 0; i < CountPage; i++)
            {
                if (i > 0)
                {
                    xmlDocument.Load(GetXMLFromRedmine(query + "&offset=" + (i * Limit)));
                }
                Result = SaveEntriesToDB(xmlDocument);
                if (Result)
                {
                    CountSavePage++;
                }

            }
            if (CountSavePage == CountPage)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private static string GetValueFromXMLNode(XmlNode xnode,string Field)
        {
            string Val = "";
            switch (Field.ToLower())
            {
                case "idporasket":
                    Val = xnode.SelectSingleNode("custom_fields/custom_field[@id=5]/value").InnerText  ;
                    break;
                case "storypoints":
                    Val = xnode.SelectSingleNode("custom_fields/custom_field[@id=9]/value").InnerText ;
                    break;
                default:
                    if (Field.Contains("."))
                    {
                        if (xnode.SelectNodes(Field.Split('.')[0]).Count > 0)
                        {
                            Val = xnode.SelectSingleNode(Field.Split('.')[0]).Attributes.GetNamedItem(Field.Split('.')[1]).InnerText;
                        }

                    }
                    else
                    {
                        Val = xnode.SelectSingleNode(Field).InnerText ;
                    }
                    break;
            }
            return Val;
        }
        private static string BuildValue(String Value, string Field) {
            var Fields = new List<string>();
            string Val = "";
            Fields.AddRange (new string[] { "id","idporasket","storypoints"});
            if (Fields.Contains(Field.ToLower()))
            {
                if (Value.Length == 0)
                {
                    Val = "0";
                }
                else {
                    Val = Value;
                }
            }
            else
            {
                Val = "'"+ Value + "'";
            }
            return Val;
        }
        private static string BuildValueString(XmlNode xnode, string Fields)
        {
            string Insert = "(";
            
            foreach (string Field in Fields.Split(','))
            {
                Insert += BuildValue(GetValueFromXMLNode(xnode, Field) ,Field)+ ",";
            }
            Insert = Insert.Remove(Insert.Length - 1, 1) + ")  ";
            return Insert;
        }
        private static string BuildUpdateString(ref ADODB.Recordset rs, XmlNode xnode, string Fields)
        {
            string Update = "Update " + Table + " SET ";
            string Val = "";
            bool BuildString = false;
            foreach (string Field in Fields.Split(','))
            {
                if (Field != "id")
                {
                    Val = GetValueFromXMLNode(xnode, Field);
                    if (Convert.ToString(rs.Fields[Field.Replace(".", "").Trim()].Value) != Val)
                    {
                        Update += Field.Replace(".", "") + "=" + BuildValue(Val, Field) + "," ;                        
                        BuildString = true;
                    }
                }
            }
            if (BuildString)
            {
                Update = Update.Remove(Update.Length - 1, 1) ;
                Update += " where id = " + xnode.SelectSingleNode("id").InnerText + " ";
            }
            else
            {
                Update = "";
            }

            return Update;
        }
        public virtual bool SaveEntriesToDB(XmlDocument xmlDocument)
        {
            string SqlInsert = "insert into  " + Table + "(" + Field.Replace(".","") +") values ";
            string SqlUpdate = "";
            bool FlagRunQuery = true;
            string SqlSelect = "Select " + Field.Replace(".", "")  + " from  " + Table + " order by id";
            DB db = DB.Instance;
            ADODB.Recordset rs = DB.Select(SqlSelect);
            if (xmlDocument.ChildNodes.Count > 0)
            {
                XmlElement xRoot = xmlDocument.DocumentElement;
                int i = 1;                
                foreach (XmlNode xnode in xRoot)
                {
                    rs.Filter = "Id = " + xnode.ChildNodes[0].InnerText;
                    if (rs.RecordCount == 0)
                    {
                        if (i > 1)
                        {
                            SqlInsert += "," ;
                        }
                        SqlInsert += BuildValueString(xnode,   Field);
                        i++;
                    }
                    else
                    {
                        SqlUpdate += BuildUpdateString(ref rs, xnode, Field);                       
                    }
                }
                if ((SqlUpdate.Length == 0) && (i == 1))
                {
                    FlagRunQuery = true;
                }
                else {
                    if (i > 1)
                    {
                        if (!(DB.ExecuteQuery(SqlInsert)))
                        {
                            FlagRunQuery = false;
                        }
                    }
                    if (SqlUpdate.Length  > 0)
                    {
                        if (!(DB.ExecuteQuery(SqlUpdate)))
                        {
                            FlagRunQuery = false;
                        }                       
                    }
                }
                return FlagRunQuery;
            }
            else
            {
                return true;
            }
        }
    }
    class Redmine : RedmineEntity 
    {
        
        public Redmine()
        {
            DB db = DB.Instance;
            GetDataFromRedmine();            
        }
        public void GetDataFromRedmine()
        {
            WriteLog.writeStringInLog ("Попытка выполнить обновление");
            if ((SaveUsers()) && (ConnectUsersFromRedmineWithAsket()))
            {
                WriteLog.writeStringInLog("Обновили пользователей.");
            }
            else
            {
                WriteLog.writeStringInLog("Во время обновления пользователей произошли ошибки.");
            }
            if (SaveIssues())
            {
                WriteLog.writeStringInLog("Обновили задачи.");
            }
            else
            {
                WriteLog.writeStringInLog("Во время обновления задач произошли ошибки.");
            }
            if ((SaveTimeEntries()) && (MarkDeletedTimeEntries()) && (InsertIntoAsketPorTimeEntries()) )
            {
                WriteLog.writeStringInLog("Обновили трудозатраты.");
            }
            else
            {
                WriteLog.writeStringInLog("Во время обновления трудозатрат произошли ошибки.");
            }
            if (DeleteAsketPor())
            {
                WriteLog.writeStringInLog("Удалили трудозатраты из Аскета.");
            }
              else
            {
                WriteLog.writeStringInLog("Во время удаления трудозатрат из Аскета произошли ошибки.");
            }
            if (SetSheduledHours())
            {
                WriteLog.writeStringInLog("Обновили планые часы по заявкам в Аскете.");
            }
            else
            {
                WriteLog.writeStringInLog("Во время обновления плановых часов по заявкам в Аскете произошли ошибки.");
            }
        }

        public override bool ReadPageFromRedmine(string Page)
        {           
            if (Page == "issues.xml")
            {
                Page += "?status_id=*&cf_5=>%3D0&";
            }
            else {
                Page +=  "?";
            }

            if (base.ReadPageFromRedmine(Page))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private  bool SaveTimeEntries()
        {
            Table = DB.TABLE_REDMINE_TIME_ENTRIES;
            Field = "id,issue.id,user.id,user.name,activity.id,activity.name,hours,comments,spent_on,created_on,updated_on";
            if (ReadPageFromRedmine("time_entries.xml"))  
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private bool SaveUsers()
        {
            Table = DB.TABLE_REDMINE_USERS;
            Field = "id, login,firstname,lastname,mail";
            if (ReadPageFromRedmine("users.xml"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private bool SaveIssues()
        {
            Table = DB.TABLE_REDMINE_ISSUES;
            Field = "id,IdPorAsket,StoryPoints";
            if (ReadPageFromRedmine("issues.xml"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
       

        private static bool MarkDeletedTimeEntries()
        {
            XmlDocument xmlDocument = new XmlDocument();            
            xmlDocument.Load(GetXMLFromRedmine("time_entries.xml?limit=" + Limit.ToString()));
            int CountPage = (Convert.ToInt32(xmlDocument.DocumentElement.Attributes.GetNamedItem("total_count").InnerText) / Limit) + 1;
            string SqlUpdate = "Update  " + DB.TABLE_REDMINE_TIME_ENTRIES + " set Deleted= 1 where id not in ( ";            
            for (int i = 0; i < CountPage; i++) {
                if (i > 0)
                {
                    xmlDocument.Load(GetXMLFromRedmine("time_entries.xml?limit=100&offset=" + (i * Limit)));
                }
                if (xmlDocument.ChildNodes.Count > 0)
                {
                    XmlElement xRoot = xmlDocument.DocumentElement;
                    int j = 0;
                    foreach (XmlNode xnode in xRoot)
                    {
                        if ((j > 0) || (i>0))
                        {
                            SqlUpdate += ",";
                        }
                        SqlUpdate += xnode.ChildNodes[0].InnerText;
                        j++;
                    }
                }
                SqlUpdate += "\n";
            }
            SqlUpdate += ")";
            if (DB.ExecuteQuery(SqlUpdate))
            {
                return true;
            }
            else
            {
                return false;
            }

        }
        private static bool DeleteAsketPor()
        {
            string Select = "Select Time.Id, IdRazrabAsket from " + DB.TABLE_REDMINE_ISSUES + " as ISSUES "  +  
                            " left join " + DB.TABLE_REDMINE_TIME_ENTRIES + " as Time on ISSUES.Id = Time.IssueId " + 
                            "  where Deleted = 1 and DeletedInAsketPor = 0";
            ADODB.Recordset RsPor = DB.Select(Select);
            if (RsPor == null)
            {
                return false;
            }
            string Update = "";
            while (!(RsPor.EOF))
            {
                if (RsPor.Fields["IdRazrabAsket"].Value > 0)
                {
                    Update = "Delete from " + DB.TABLE_REPORT_DEVELOPER + " where Код=" + RsPor.Fields["IdRazrabAsket"].Value;
                    DB.ExecuteQuery(Update);
                    Update = "Update" + DB.TABLE_REDMINE_TIME_ENTRIES + " set DeletedInAsketPor= 1  where id=" + RsPor.Fields["Id"].Value;
                    DB.ExecuteQuery(Update);
                }
                RsPor.MoveNext();
            }
            RsPor.Close();
            return true;
        }              
        private static bool ConnectUsersFromRedmineWithAsket()
        {            
            string SqlSelect = "select Код,email from " + DB.TABLE_ASKET_USERS + " where уволен = 0 and email in( select mail from " + DB.TABLE_REDMINE_USERS + " where IdUserFromFam = 0)";
            ADODB.Recordset rs = DB.Select(SqlSelect);
            ADODB.Recordset users = DB.Update("Select  * from "  + DB.TABLE_REDMINE_USERS );
            while (!rs.EOF)
            {
                users.Filter = "mail = '" + rs.Fields["email"].Value + "'";
                users.Fields["IdUserFromFam"].Value = rs.Fields["Код"].Value;
                rs.MoveNext();
            }
            users.Update();
            users.Close();

            rs.Close();
            rs = null;
            return true;
        }        
        private static bool InsertIntoAsketPorTimeEntries()
        {
            string Select = "select з.IdPorAsket, тз.IssueId, ТЗ.Spent_On,тз.Hours, Ф.UserName ,  тз.Comments + ' Импортировано из Redmine. ' as Comments, тз.id" +
                            " from      " + DB.TABLE_REDMINE_TIME_ENTRIES + " as ТЗ " +
                            " left join " + DB.TABLE_REDMINE_USERS + "  as П on тз.UserId = п.id " +
                            " left join " + DB.TABLE_ASKET_USERS + " as Ф On п.IdUserFromFam = Ф.код " +
                            " left join " + DB.TABLE_REDMINE_ISSUES + "  as З on ТЗ.IssueId = З.id " +
                            " where Deleted = 0 and InsertIntoAsketPor = 0 and з.idPorAsket is not null ";
            ADODB.Recordset rsSelect = DB.Select(Select);
            if (rsSelect == null)
            {
                return false;
            }
            ADODB.Recordset RsAddPor = DB.Update("select top 1 Код, КодПоручения, Дата, Часов, Кто, Примечания from " + DB.TABLE_REPORT_DEVELOPER);
            if (RsAddPor == null)
            {
                return false;
            }
            ADODB.Recordset RsUpdateTimeEntries = DB.Update("Select Id, InsertIntoAsketPor,IdRazrabAsket from  " + DB.TABLE_REDMINE_TIME_ENTRIES + " where Deleted = 0 and InsertIntoAsketPor = 0 and IdRazrabAsket =0");
            if (RsUpdateTimeEntries == null)
            {
                return false;
            }
            while (!(rsSelect.EOF)) {
                RsAddPor.AddNew();
                RsAddPor.Fields["КодПоручения"].Value = rsSelect.Fields["IdPorAsket"].Value;
                RsAddPor.Fields["Дата"].Value = rsSelect.Fields["Spent_On"].Value;
                RsAddPor.Fields["Часов"].Value = rsSelect.Fields["Hours"].Value;
                RsAddPor.Fields["Кто"].Value = rsSelect.Fields["UserName"].Value;
                RsAddPor.Fields["Примечания"].Value = rsSelect.Fields["Comments"].Value;
                RsAddPor.Update();
                RsUpdateTimeEntries.Filter = "id = " + rsSelect.Fields["id"].Value;
                if (RsUpdateTimeEntries.RecordCount == 1)
                {
                    RsUpdateTimeEntries.Fields["InsertIntoAsketPor"].Value = true;
                    RsUpdateTimeEntries.Fields["IdRazrabAsket"].Value = RsAddPor.Fields["Код"].Value;
                    RsUpdateTimeEntries.Update();
                }
                rsSelect.MoveNext();                

             }
            rsSelect.Close();
            RsAddPor.Close();            
            return true;
        }
        private static bool SetSheduledHours()
        {
            string SqlSelect = " select п.Код, П.ВремяПлан,  isnull(П.ВремяРазработкиПлан,0) ВремяРазработкиПлан , " +
                              " cast(Format(GetDate() ,'dd.MM.yyyy') as varchar(20)) + ' Обновление планируемых часов. Было - '+   cast(isnull(П.ВремяРазработкиПлан,0) as varchar(10)) + ' Стало ' + cast(cast(sp.Hours as int) as varchar(3)) + char(10)  + char(13) + " + Environment.NewLine +
                                " + replace(ХодВыполнения,'''','''''')    as ХодВыполнения, sp.Hours  " +
                               " from stmain.dbo.АскетПоручения  as П left join(" +
                                                                              " SELECT IdPorAsket, (sum(StoryPoints) * 8) Hours " +
                                                                              " FROM  " + DB.TABLE_REDMINE_ISSUES +
                                                                              " group by IdPorAsket " +
                                                                               ") as SP on п.код = sp.IdPorAsket " +
                               " where  sp.Hours <> isnull(П.ВремяРазработкиПлан,0)   and  sp.Hours >0";
            ADODB.Recordset rs = DB.Select(SqlSelect);
            string SqlUpdate = "";
            while (!rs.EOF)
            {
                SqlUpdate = "Update stmain.dbo.АскетПоручения set ВремяРазработкиПлан =" + Convert.ToString(Convert.ToInt16 (rs.Fields["Hours"].Value)).Replace(',','.') + ", ХодВыполнения= '" + rs.Fields["ХодВыполнения"].Value + "' where Код = " + rs.Fields["Код"].Value;
                DB.ExecuteQuery(SqlUpdate);
                rs.MoveNext();
            }
            return true;
        }
    }
}