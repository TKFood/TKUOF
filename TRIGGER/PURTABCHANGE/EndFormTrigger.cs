using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Configuration;
using System.Data.Common;
using System.Data.Odbc;
using System.Data.OleDb;
using Ede.Uof.Utility.Data;
using Ede.Uof.WKF.ExternalUtility;
using System.Xml;
using System.Xml.Linq;

namespace TKUOF.TRIGGER.PURTABCHANGE
{
    //請購單變更的核準

    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {
            
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string FORMID = null;
            string TA001=null;
            string TA002 = null;
            string TA006 = null;
            string TB003 = null;
            string TB004 = null;
            string TB009 = null;
            string TB011 = null;
            string TB012 = null;
            string ADDSQL = null;




            ////找到明細的XML物件
            //XElement xe = XElement.Parse(applyTask.CurrentDocXML);
            //var nodes = (from xl in xe.Element("FormFieldValue").Elements("FieldItem")
            //             where xl.Attribute("fieldId").Value == "TB"
            //             select xl).FirstOrDefault();

            ////找出PO_DETAIL的所有ROW集合
            //foreach (var node in nodes.Element("DataGrid").Elements("Row"))
            //{
            //    //找出所有ROW的CELL
            //    var cells = node.Elements("Cell");

            //    var PO_ITEM01的值 = cells.Where(p => p.Attribute("fieldId").Value == "PO_ITEM01").FirstOrDefault().Attribute("fieldValue");
            //    var MATERIAL_011的值 = cells.Where(p => p.Attribute("fieldId").Value == "MATERIAL_01").FirstOrDefault().Attribute("fieldValue");
            //    //........以此類推
            //}

            ///核準
            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(applyTask.CurrentDocXML);
                FORMID = applyTask.FormNumber;
                TA001 = applyTask.Task.CurrentDocument.Fields["TA001"].FieldValue.ToString().Trim();
                TA002 = applyTask.Task.CurrentDocument.Fields["TA002"].FieldValue.ToString().Trim();
                TA006 = applyTask.Task.CurrentDocument.Fields["TA006"].FieldValue.ToString().Trim();


                //針對DETAIL抓出來的資料作處理
                Ede.Uof.WKF.Design.FieldDataGrid grid = applyTask.Task.CurrentDocument.Fields["TB"] as Ede.Uof.WKF.Design.FieldDataGrid;
                foreach (var row in grid.FieldDataGridValue.RowValueList)
                {
                    foreach (var cell in row.CellValueList)
                    {
                        if (cell.fieldId == "TB003")
                        {
                            TB003 = cell.fieldValue;
                        }
                        if (cell.fieldId == "TB004")
                        {
                            TB004 = cell.fieldValue;
                        }
                        if (cell.fieldId == "TB009")
                        {
                            TB009 = cell.fieldValue;
                        }
                        if (cell.fieldId == "TB011")
                        {
                            TB011 = cell.fieldValue;
                        }
                        if (cell.fieldId == "TB012")
                        {
                            TB012 = cell.fieldValue;
                        }
                    }

                    if (!string.IsNullOrEmpty(FORMID) && !string.IsNullOrEmpty(TA001) && !string.IsNullOrEmpty(TA002))
                    {
                        ADDSQL = ADDSQL+SETPURTABUOFCHANGE(FORMID, TA001, TA002, TA006, TB003, TB004, TB009, TB011, TB012);
                        ADDSQL = ADDSQL + " ";
                    }
                }

                ADDPURTABUOFCHANGE(FORMID,TA001,TA002, ADDSQL);
            }
       

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public string SETPURTABUOFCHANGE(string FORMID, string TA001, string TA002, string TA006, string TB003, string TB004, string TB009, string TB011, string TB012)
        {
            StringBuilder SQL = new StringBuilder();
            SQL.AppendFormat(@" 
                                INSERT INTO [TKPUR].[dbo].[PURTABUOFCHANGE]
                                ([FORMID],[TA001],[TA002],[TA006],[TB003],[TB004],[TB009],[TB011],[TB012])
                                VALUES
                                (@FORMID,'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')

                                ",  TA001, TA002, TA006, TB003, TB004, TB009, TB011, TB012);

            return SQL.ToString();
        }

        public void ADDPURTABUOFCHANGE(string FORMID,string TA001,string TA002,string ADDSQL)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            //StringBuilder queryString = new StringBuilder();
            //queryString.AppendFormat(@"
                                       
            //                            ", FORMID);

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(ADDSQL.ToString(), connection);
                    command.Parameters.Add("@FORMID", SqlDbType.NVarChar).Value = FORMID;
         
                    command.Connection.Open();

                    int count = command.ExecuteNonQuery();

                    connection.Close();
                    connection.Dispose();

                    UPDATEPURTATB(FORMID, TA001, TA002);
                }
            }
            catch
            {

            }
            finally
            {
               
            }
                
        }

        public void UPDATEPURTATB(string FORMID, string TA001, string TA002)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                        UPDATE [TK].[dbo].[PURTA]
                                        SET [PURTA].[TA006]=[PURTABUOFCHANGE].[TA006]
                                        FROM [TKPUR].[dbo].[PURTABUOFCHANGE]
                                        WHERE [PURTA].TA001=@TA001 AND [PURTA].TA002=@TA002
                                        AND [PURTABUOFCHANGE].FORMID=@FORMID

                                        UPDATE [TK].[dbo].[PURTB]
                                        SET [PURTB].[TB004]=[PURTABUOFCHANGE].[TB004],[PURTB].[TB009]=[PURTABUOFCHANGE].[TB009],[PURTB].[TB011]=[PURTABUOFCHANGE].[TB011],[PURTB].[TB012]=[PURTABUOFCHANGE].[TB012]
                                        ,[PURTB].[TB005]=INVMB.MB002
                                        ,[PURTB].[TB006]=INVMB.MB003
                                        ,[PURTB].[TB007]=INVMB.MB004
                                        ,[PURTB].[TB017]=INVMB.MB050 
                                        ,[PURTB].[TB018]=(MB050*[PURTABUOFCHANGE].[TB009]) 
                                        FROM [TKPUR].[dbo].[PURTABUOFCHANGE],[TK].dbo.INVMB
                                        WHERE [PURTABUOFCHANGE].TB004=INVMB.MB001
                                        AND [PURTB].TB003=[PURTABUOFCHANGE].TB003
                                        AND [PURTB].TB001=@TA001 AND [PURTB].TB002=@TA002
                                        AND [PURTABUOFCHANGE].FORMID=@FORMID

                                        INSERT INTO [TK].[dbo].[PURTB]
                                        (
                                        [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]
                                        ,[TB001],[TB002],[TB003],[TB004],[TB005],[TB006],[TB007],[TB008],[TB009],[TB010]
                                        ,[TB011],[TB012],[TB013],[TB014],[TB015],[TB016],[TB017],[TB018],[TB019],[TB020]
                                        ,[TB021],[TB022],[TB023],[TB024],[TB025],[TB026],[TB027],[TB028],[TB029],[TB030]
                                        ,[TB031],[TB032],[TB033],[TB034],[TB035],[TB036],[TB037],[TB038],[TB039],[TB040]
                                        ,[TB041],[TB042],[TB043],[TB044],[TB045],[TB046],[TB047],[TB048],[TB049],[TB050]
                                        ,[TB051],[TB052],[TB053],[TB054],[TB055],[TB056],[TB057],[TB058],[TB059],[TB060]
                                        ,[TB061],[TB062],[TB063],[TB064],[TB065],[TB066],[TB067],[TB068],[TB069],[TB070]
                                        ,[TB071],[TB072],[TB073],[TB074],[TB075],[TB076],[TB077],[TB078],[TB079],[TB080]
                                        ,[TB081],[TB082],[TB083],[TB084],[TB085],[TB086],[TB087],[TB088],[TB089],[TB090]
                                        ,[TB091],[TB092],[TB093],[TB094],[TB095],[TB096],[TB097],[TB098],[TB099]
                                        ,[UDF01],[UDF02],[UDF03],[UDF04],[UDF05],[UDF06],[UDF07],[UDF08],[UDF09],[UDF10]
                                        )
                                        SELECT [PURTB].[COMPANY],[PURTB].[CREATOR],[PURTB].[USR_GROUP],[PURTB].[CREATE_DATE],[PURTB].[MODIFIER],[PURTB].[MODI_DATE],[PURTB].[FLAG],[PURTB].[CREATE_TIME],[PURTB].[MODI_TIME],[PURTB].[TRANS_TYPE],[PURTB].[TRANS_NAME],[PURTB].[sync_date],[PURTB].[sync_time],[PURTB].[sync_mark],[PURTB].[sync_count],[PURTB].[DataUser],[PURTB].[DataGroup]
                                        ,[TB001],[TB002],[PURTABUOFCHANGE].[TB003] TB003,[PURTABUOFCHANGE].[TB004] TB004,INVMB.MB002 [TB005],INVMB.MB003 [TB006],INVMB.MB004 [TB007],[TB008],[PURTABUOFCHANGE].[TB009] TB009,MB032 [TB010]
                                        ,[PURTABUOFCHANGE].[TB011] TB011,[PURTABUOFCHANGE].[TB012] TB012,[TB013],[TB014],[TB015],[TB016],MB050 [TB017],(MB050*[PURTABUOFCHANGE].[TB009]) [TB018], [TB019],[TB020]
                                        ,[TB021],[TB022],[TB023],[TB024],'Y' [TB025],[TB026],[TB027],[TB028],[TB029],[TB030]
                                        ,[TB031],[TB032],[TB033],[TB034],[TB035],[TB036],[TB037],[TB038],[TB039],[TB040]
                                        ,[TB041],[TB042],[TB043],[TB044],[TB045],[TB046],[TB047],[TB048],[TB049],[TB050]
                                        ,[TB051],[TB052],[TB053],[TB054],[TB055],[TB056],[TB057],[TB058],[TB059],[TB060]
                                        ,[TB061],[TB062],[TB063],[TB064],[TB065],[TB066],[TB067],[TB068],[TB069],[TB070]
                                        ,[TB071],[TB072],[TB073],[TB074],[TB075],[TB076],[TB077],[TB078],[TB079],[TB080]
                                        ,[TB081],[TB082],[TB083],[TB084],[TB085],[TB086],[TB087],[TB088],[TB089],[TB090]
                                        ,[TB091],[TB092],[TB093],[TB094],[TB095],[TB096],[TB097],[TB098],[TB099]
                                        ,[PURTB].[UDF01],[PURTB].[UDF02],[PURTB].[UDF03],[PURTB].[UDF04],[PURTB].[UDF05],[PURTB].[UDF06],[PURTB].[UDF07],[PURTB].[UDF08],[PURTB].[UDF09],[PURTB].[UDF10]
                                        FROM [TK].[dbo].[PURTB],[TKPUR].[dbo].[PURTABUOFCHANGE],[TK].dbo.INVMB
                                        WHERE [PURTABUOFCHANGE].TA001=[PURTB].TB001 AND [PURTABUOFCHANGE].TA002=[PURTB].TB002 
                                        AND [PURTABUOFCHANGE].TB004=INVMB.MB001
                                        AND [PURTABUOFCHANGE].TB003 NOT IN (SELECT TB003 FROM [TK].[dbo].[PURTB] WHERE TB001=@TA001 AND TB002=@TA002)
                                        AND [PURTB].TB001=@TA001 AND [PURTB].TB002=@TA002
                                        AND [PURTABUOFCHANGE].FORMID=@FORMID

                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    
                    command.Parameters.Add("@FORMID", SqlDbType.NVarChar).Value = FORMID;
                    command.Parameters.Add("@TA001", SqlDbType.NVarChar).Value = TA001;
                    command.Parameters.Add("@TA002", SqlDbType.NVarChar).Value = TA002;

                    command.Connection.Open();

                    int count = command.ExecuteNonQuery();

                    connection.Close();
                    connection.Dispose();

                }
            }
            catch
            {

            }
            finally
            {

            }

        }

    }
}
