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
            string VERSIONS = null;

            SqlConnection sqlConn = new SqlConnection();
            string connectionString;
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();           
            int result;


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
                VERSIONS = applyTask.Task.CurrentDocument.Fields["VERSIONS"].FieldValue.ToString().Trim();


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
                        ADDSQL = ADDSQL+ SETPURTATBUOFCHANGE(FORMID, TA001, TA002, TA006, TB003, TB004, TB009, TB011, TB012);
                        ADDSQL = ADDSQL + " ";
                    }
                }

                ADDPURTATBUOFCHANGE(FORMID,TA001,TA002, ADDSQL);

                NEWPURTEPURTF(TA001, TA002,VERSIONS);

            }

       

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public string SETPURTATBUOFCHANGE(string FORMID, string TA001, string TA002, string TA006, string TB003, string TB004, string TB009, string TB011, string TB012)
        {
            StringBuilder SQL = new StringBuilder();
            SQL.AppendFormat(@" 
                                INSERT INTO [TKPUR].[dbo].[PURTATBUOFCHANGE]
                                ([FORMID],[TA001],[TA002],[TA006],[TB003],[TB004],[TB009],[TB011],[TB012])
                                VALUES
                                (@FORMID,'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')

                                ",  TA001, TA002, TA006, TB003, TB004, TB009, TB011, TB012);

            return SQL.ToString();
        }

        public void ADDPURTATBUOFCHANGE(string FORMID,string TA001,string TA002,string ADDSQL)
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
                                        SET [PURTA].[TA006]=[PURTATBUOFCHANGE].[TA006], [PURTA].[UDF04]=@FORMID
                                        FROM [TKPUR].[dbo].[PURTATBUOFCHANGE]
                                        WHERE [PURTA].TA001=@TA001 AND [PURTA].TA002=@TA002
                                        AND [PURTATBUOFCHANGE].FORMID=@FORMID

                                        UPDATE [TK].[dbo].[PURTB]
                                        SET [PURTB].[TB004]=[PURTATBUOFCHANGE].[TB004],[PURTB].[TB009]=[PURTATBUOFCHANGE].[TB009],[PURTB].[TB011]=[PURTATBUOFCHANGE].[TB011],[PURTB].[TB012]=[PURTATBUOFCHANGE].[TB012]
                                        ,[PURTB].[TB005]=INVMB.MB002
                                        ,[PURTB].[TB006]=INVMB.MB003
                                        ,[PURTB].[TB007]=INVMB.MB004
                                        ,[PURTB].[TB017]=INVMB.MB050 
                                        ,[PURTB].[TB018]=(MB050*[PURTATBUOFCHANGE].[TB009]) 
                                        ,[PURTB].[TB021]='N'
                                        FROM [TKPUR].[dbo].[PURTATBUOFCHANGE],[TK].dbo.INVMB
                                        WHERE [PURTATBUOFCHANGE].TB004=INVMB.MB001
                                        AND [PURTB].TB003=[PURTATBUOFCHANGE].TB003
                                        AND [PURTB].TB001=@TA001 AND [PURTB].TB002=@TA002
                                        AND [PURTATBUOFCHANGE].FORMID=@FORMID

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
                                        ,[TB001],[TB002],[PURTATBUOFCHANGE].[TB003] TB003,[PURTATBUOFCHANGE].[TB004] TB004,INVMB.MB002 [TB005],INVMB.MB003 [TB006],INVMB.MB004 [TB007],[TB008],[PURTATBUOFCHANGE].[TB009] TB009,MB032 [TB010]
                                        ,[PURTATBUOFCHANGE].[TB011] TB011,[PURTATBUOFCHANGE].[TB012] TB012,[TB013],[TB014],[TB015],[TB016],MB050 [TB017],(MB050*[PURTATBUOFCHANGE].[TB009]) [TB018], [TB019],[TB020]
                                        ,[TB021],[TB022],[TB023],[TB024],'Y' [TB025],[TB026],[TB027],[TB028],[TB029],[TB030]
                                        ,[TB031],[TB032],[TB033],[TB034],[TB035],[TB036],[TB037],[TB038],[TB039],[TB040]
                                        ,[TB041],[TB042],[TB043],[TB044],[TB045],[TB046],[TB047],[TB048],[TB049],[TB050]
                                        ,[TB051],[TB052],[TB053],[TB054],[TB055],[TB056],[TB057],[TB058],[TB059],[TB060]
                                        ,[TB061],[TB062],[TB063],[TB064],[TB065],[TB066],[TB067],[TB068],[TB069],[TB070]
                                        ,[TB071],[TB072],[TB073],[TB074],[TB075],[TB076],[TB077],[TB078],[TB079],[TB080]
                                        ,[TB081],[TB082],[TB083],[TB084],[TB085],[TB086],[TB087],[TB088],[TB089],[TB090]
                                        ,[TB091],[TB092],[TB093],[TB094],[TB095],[TB096],[TB097],[TB098],[TB099]
                                        ,[PURTB].[UDF01],[PURTB].[UDF02],[PURTB].[UDF03],[PURTB].[UDF04],[PURTB].[UDF05],[PURTB].[UDF06],[PURTB].[UDF07],[PURTB].[UDF08],[PURTB].[UDF09],[PURTB].[UDF10]
                                        FROM [TK].[dbo].[PURTB],[TKPUR].[dbo].[PURTATBUOFCHANGE],[TK].dbo.INVMB
                                        WHERE [PURTATBUOFCHANGE].TA001=[PURTB].TB001 AND [PURTATBUOFCHANGE].TA002=[PURTB].TB002  AND [PURTB].TB003=(SELECT TOP 1 TB003 FROM [TK].[dbo].[PURTB] WHERE  [PURTB].TB001=@TA001 AND [PURTB].TB002=@TA002)
                                        AND [PURTATBUOFCHANGE].TB004=INVMB.MB001
                                        AND [PURTATBUOFCHANGE].TB003 NOT IN (SELECT TB003 FROM [TK].[dbo].[PURTB] WHERE TB001=@TA001 AND TB002=@TA002)
                                        AND [PURTB].TB001=@TA001 AND [PURTB].TB002=@TA002
                                        AND [PURTATBUOFCHANGE].FORMID=@FORMID

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

        public void NEWPURTEPURTF(string TA001, string TA002, string VERSIONS)
        {

            //A311 20221101011 1
            //檢查請購變更單的採購單，是否有採購變更單未核準
            DataTable DTCHECKPURTEPURTF = CHECKPURTEPURTF(TA001, TA002, VERSIONS);

            if (DTCHECKPURTEPURTF == null)
            {
                //找出請購變更單有幾張採購單，要1對多
                DataTable DTPURTCPURTD = SEARCHPURTCPURTD(TA001, TA002, VERSIONS);
                //DataTable DTPURTCPURTD = SEARCHPURTCPURTD("A312", "20221116001", "2");
                DataTable DTOURTE = new DataTable();

                //找出採購單跟最大的版次
                if (DTPURTCPURTD.Rows.Count > 0)
                {
                    DTOURTE = FINDPURTE(DTPURTCPURTD);
                }

                //新增採購變更單
                if (DTOURTE.Rows.Count > 0)
                {
                    ADDTOPURTEPURTF(DTOURTE);
                }
            }
            else
            {
                //StringBuilder MESS = new StringBuilder();
                //foreach (DataRow DR in DTCHECKPURTEPURTF.Rows)
                //{
                //    MESS.AppendFormat(@" 採購變更單:" + DR["TE001"].ToString() + " " + DR["TE002"].ToString() + "" + "變更版次:" + DR["TE003"].ToString() + " 沒有核準 ");
                //}

                //MessageBox.Show(MESS.ToString());
            }


        }

        public DataTable CHECKPURTEPURTF(string TA001, string TA002, string VERSIONS)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

          

            try
            {
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ConnectionString);
                SqlConnection  sqlConn = new SqlConnection(sqlsb.ConnectionString);
                StringBuilder sbSql = new StringBuilder();
                StringBuilder sbSqlQuery = new StringBuilder();

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT TE001,TE002,TE003
                                    FROM [TK].dbo.PURTE
                                    WHERE TE017 IN ('N')
                                    AND TE001+TE002 IN 
                                    (
                                    SELECT TD001+TD002
                                    FROM [TK].dbo.PURTD
                                    WHERE TD026+TD027+TD028 IN 
                                    (
                                    SELECT TA001+TA002+TB003
                                    FROM [TKPUR].[dbo].[PURTATBCHAGE]
                                    WHERE  TA001='{0}' AND TA002='{1}' AND VERSIONS='{2}'
                                    )
                                    GROUP BY  TD001,TD002
                                    )
                                    ", TA001, TA002, VERSIONS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {

                    return ds.Tables["TEMPds1"];
                }
                else
                {
                    return null;

                }


            }
            catch
            {
                return null;
            }
            finally
            {

            }



        }

        public DataTable SEARCHPURTCPURTD(string TA001, string TA002, string VERSIONS)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            try
            {

                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ConnectionString);
                SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString);
                StringBuilder sbSql = new StringBuilder();
                StringBuilder sbSqlQuery = new StringBuilder();


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                  
                                    SELECT TD001,TD002,'{0}' TA001,'{1}' TA002,'{2}' VERSIONS
                                    FROM [TK].dbo.PURTD
                                    WHERE TD026+TD027+TD028 IN 
                                    (
                                    SELECT TA001+TA002+TB003
                                    FROM [TKPUR].[dbo].[PURTATBCHAGE]
                                    WHERE  TA001='{0}' AND TA002='{1}' AND VERSIONS='{2}'
                                    )
                                    AND TD018 IN ('Y')
                                    GROUP BY  TD001,TD002
                                    ", TA001, TA002, VERSIONS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {

                    return ds.Tables["TEMPds1"];
                }
                else
                {
                    return null;

                }


            }
            catch
            {
                return null;
            }
            finally
            {

            }



        }

        public DataTable FINDPURTE(DataTable DTTEMP)
        {
            DataTable DT = new DataTable();
            DT.Clear();
            DT.Columns.Add("TE001");
            DT.Columns.Add("TE002");
            DT.Columns.Add("TE003");
            DT.Columns.Add("TA001");
            DT.Columns.Add("TA002");
            DT.Columns.Add("VERSIONS");



            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            string TE001 = null;
            string TE002 = null;
            string TA001 = null;
            string TA002 = null;
            string VERSIONS = null;

            if (DTTEMP.Rows.Count > 0)
            {
                foreach (DataRow DR in DTTEMP.Rows)
                {

                    TE001 = DR["TD001"].ToString();
                    TE002 = DR["TD002"].ToString();
                    TA001 = DR["TA001"].ToString();
                    TA002 = DR["TA002"].ToString();
                    VERSIONS = DR["VERSIONS"].ToString();

                    try
                    {
                        SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ConnectionString);
                        SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString);
                        StringBuilder sbSql = new StringBuilder();
                        StringBuilder sbSqlQuery = new StringBuilder();



                        sbSql.Clear();
                        sbSqlQuery.Clear();


                        sbSql.AppendFormat(@"  
                                            SELECT TOP 1 TE001,TE002,TE003
                                            FROM [TK].dbo.PURTE
                                            WHERE TE001='{0}' AND TE002='{1}'
                                            ORDER BY TE001 DESC,TE002 DESC,TE003 DESC
                                             ", TE001, TE002);

                        adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                        sqlCmdBuilder = new SqlCommandBuilder(adapter);
                        sqlConn.Open();
                        ds.Clear();
                        adapter.Fill(ds, "TEMPds1");
                        sqlConn.Close();


                        if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                        {
                            int serno = Convert.ToInt16(ds.Tables["TEMPds1"].Rows[0]["TE003"].ToString());
                            serno = serno + 1;
                            string temp = serno.ToString();
                            temp = temp.PadLeft(4, '0');

                            DataRow NEWDR = DT.NewRow();
                            NEWDR["TE001"] = TE001;
                            NEWDR["TE002"] = TE002;
                            NEWDR["TE003"] = temp;
                            NEWDR["TA001"] = TA001;
                            NEWDR["TA002"] = TA002;
                            NEWDR["VERSIONS"] = VERSIONS;
                            DT.Rows.Add(NEWDR);

                        }
                        else
                        {
                            DataRow NEWDR = DT.NewRow();
                            NEWDR["TE001"] = TE001;
                            NEWDR["TE002"] = TE002;
                            NEWDR["TE003"] = "0001";
                            NEWDR["TA001"] = TA001;
                            NEWDR["TA002"] = TA002;
                            NEWDR["VERSIONS"] = VERSIONS;
                            DT.Rows.Add(NEWDR);

                        }


                    }
                    catch
                    {
                        return null;
                    }
                    finally
                    {

                    }
                }

                return DT;
            }
            else
            {
                return null;
            }
        }


        public void ADDTOPURTEPURTF(DataTable NEWPURTEPURTF)
        {
            SqlCommand cmd = new SqlCommand();

            if (NEWPURTEPURTF.Rows.Count > 0)
            {
                try
                {
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ConnectionString);
                    SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString);
                    StringBuilder sbSql = new StringBuilder();
                    StringBuilder sbSqlQuery = new StringBuilder();


                    sqlConn.Close();
                    sqlConn.Open();
                    SqlTransaction tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    foreach (DataRow DR in NEWPURTEPURTF.Rows)
                    {
                        sbSql.AppendFormat(@"  
                                   
                                            INSERT INTO [TK].[dbo].[PURTF]
                                            (
                                            [COMPANY]
                                            ,[CREATOR]
                                            ,[USR_GROUP]
                                            ,[CREATE_DATE]
                                            ,[MODIFIER]
                                            ,[MODI_DATE]
                                            ,[FLAG]
                                            ,[CREATE_TIME]
                                            ,[MODI_TIME]
                                            ,[TRANS_TYPE]
                                            ,[TRANS_NAME]
                                            ,[sync_date]
                                            ,[sync_time]
                                            ,[sync_mark]
                                            ,[sync_count]
                                            ,[DataUser]
                                            ,[DataGroup]
                                            ,[TF001]
                                            ,[TF002]
                                            ,[TF003]
                                            ,[TF004]
                                            ,[TF005]
                                            ,[TF006]
                                            ,[TF007]
                                            ,[TF008]
                                            ,[TF009]
                                            ,[TF010]
                                            ,[TF011]
                                            ,[TF012]
                                            ,[TF013]
                                            ,[TF014]
                                            ,[TF015]
                                            ,[TF016]
                                            ,[TF017]
                                            ,[TF018]
                                            ,[TF019]
                                            ,[TF020]
                                            ,[TF021]
                                            ,[TF022]
                                            ,[TF023]
                                            ,[TF024]
                                            ,[TF025]
                                            ,[TF026]
                                            ,[TF027]
                                            ,[TF028]
                                            ,[TF029]
                                            ,[TF030]
                                            ,[TF031]
                                            ,[TF032]
                                            ,[TF033]
                                            ,[TF034]
                                            ,[TF035]
                                            ,[TF036]
                                            ,[TF037]
                                            ,[TF038]
                                            ,[TF039]
                                            ,[TF040]
                                            ,[TF041]
                                            ,[TF104]
                                            ,[TF105]
                                            ,[TF106]
                                            ,[TF107]
                                            ,[TF108]
                                            ,[TF109]
                                            ,[TF110]
                                            ,[TF111]
                                            ,[TF112]
                                            ,[TF113]
                                            ,[TF114]
                                            ,[TF118]
                                            ,[TF119]
                                            ,[TF120]
                                            ,[TF121]
                                            ,[TF122]
                                            ,[TF123]
                                            ,[TF124]
                                            ,[TF125]
                                            ,[TF126]
                                            ,[TF127]
                                            ,[TF128]
                                            ,[TF129]
                                            ,[TF130]
                                            ,[TF131]
                                            ,[TF132]
                                            ,[TF133]
                                            ,[TF134]
                                            ,[TF135]
                                            ,[TF136]
                                            ,[TF137]
                                            ,[TF138]
                                            ,[TF139]
                                            ,[TF140]
                                            ,[TF141]
                                            ,[TF142]
                                            ,[TF143]
                                            ,[TF144]
                                            ,[TF145]
                                            ,[TF146]
                                            ,[TF147]
                                            ,[TF148]
                                            ,[TF149]
                                            ,[TF150]
                                            ,[TF151]
                                            ,[TF152]
                                            ,[TF153]
                                            ,[TF154]
                                            ,[TF155]
                                            ,[TF156]
                                            ,[TF157]
                                            ,[TF158]
                                            ,[TF159]
                                            ,[TF160]
                                            ,[TF161]
                                            ,[TF162]
                                            ,[TF163]
                                            ,[TF164]
                                            ,[TF165]
                                            ,[TF166]
                                            ,[TF167]
                                            ,[TF168]
                                            ,[TF169]
                                            ,[TF170]
                                            ,[TF171]
                                            ,[TF172]
                                            ,[TF173]
                                            ,[UDF01]
                                            ,[UDF02]
                                            ,[UDF03]
                                            ,[UDF04]
                                            ,[UDF05]
                                            ,[UDF06]
                                            ,[UDF07]
                                            ,[UDF08]
                                            ,[UDF09]
                                            ,[UDF10]
                                            )
                                            SELECT 
                                            PURTD.[COMPANY]
                                            ,PURTD.[CREATOR] AS [CREATOR]
                                            ,PURTD.[USR_GROUP] AS [USR_GROUP]
                                            ,PURTD.[CREATE_DATE] AS [CREATE_DATE]
                                            ,PURTD.[MODIFIER] AS [MODIFIER]
                                            ,PURTD.[MODI_DATE] AS [MODI_DATE]
                                            ,PURTD.[FLAG] AS [FLAG]
                                            ,PURTD.[CREATE_TIME] AS [CREATE_TIME]
                                            ,PURTD.[MODI_TIME] AS [MODI_TIME]
                                            ,PURTD.[TRANS_TYPE] AS [TRANS_TYPE]
                                            ,'PURI08' AS [TRANS_NAME]
                                            ,PURTD.[sync_date] AS [sync_date]
                                            ,PURTD.[sync_time] AS [sync_time]
                                            ,PURTD.[sync_mark] AS [sync_mark]
                                            ,PURTD.[sync_count] AS [sync_count]
                                            ,PURTD.[DataUser] AS [DataUser]
                                            ,PURTD.[DataGroup] AS [DataGroup]
                                            ,TD001 AS [TF001]
                                            ,TD002 AS [TF002]
                                            ,'{3}' AS [TF003]
                                            ,TD003 AS [TF004]
                                            ,[PURTATBCHAGE].TB004 AS [TF005]
                                            ,[PURTATBCHAGE].TB005 AS [TF006]
                                            ,[PURTATBCHAGE].TB006 AS [TF007]
                                            ,TD007 AS [TF008]
                                            ,[PURTATBCHAGE].TB009 AS [TF009]
                                            ,TD009 AS [TF010]
                                            ,TD010 AS [TF011]
                                            ,[PURTATBCHAGE].TB009*TD010 AS [TF012]
                                            ,[PURTATBCHAGE].TB011 AS [TF013]
                                            ,'N' AS [TF014]
                                            ,TD015 AS [TF015]
                                            ,'N' AS [TF016]
                                            ,[PURTATBCHAGE].TB012 AS [TF017]
                                            ,TD019 AS [TF018]
                                            ,TD020 AS [TF019]
                                            ,TD022 AS [TF020]
                                            ,TD025 AS [TF021]
                                            ,TD017 AS [TF022]
                                            ,TD029 AS [TF023]
                                            ,TD030 AS [TF024]
                                            ,TD032 AS [TF025]
                                            ,TD033 AS [TF026]
                                            ,TD036 AS [TF027]
                                            ,TD037 AS [TF028]
                                            ,TD038 AS [TF029]
                                            ,TD014 AS [TF030]
                                            ,'' AS [TF031]
                                            ,'' AS [TF032]
                                            ,'' AS [TF033]
                                            ,'' AS [TF034]
                                            ,'' AS [TF035]
                                            ,0 AS [TF036]
                                            ,0 AS [TF037]
                                            ,'' AS [TF038]
                                            ,'' AS [TF039]
                                            ,'' AS [TF040]
                                            ,0 AS [TF041]
                                            ,TD003 AS [TF104]
                                            ,TD004 AS [TF105]
                                            ,TD005 AS [TF106]
                                            ,TD006 AS [TF107]
                                            ,TD007 AS [TF108]
                                            ,TD008 AS [TF109]
                                            ,TD009 AS [TF110]
                                            ,TD010 AS [TF111]
                                            ,TD011 AS [TF112]
                                            ,TD012 AS [TF113]
                                            ,TD016 AS [TF114]
                                            ,TD019 AS [TF118]
                                            ,TD020 AS [TF119]
                                            ,TD022 AS [TF120]
                                            ,TD025 AS [TF121]
                                            ,TD017 AS [TF122]
                                            ,TD029 AS [TF123]
                                            ,TD030 AS [TF124]
                                            ,TD031 AS [TF125]
                                            ,TD032 AS [TF126]
                                            ,TD033 AS [TF127]
                                            ,TD034 AS [TF128]
                                            ,TD035 AS [TF129]
                                            ,TD034 AS [TF130]
                                            ,TD035 AS [TF131]
                                            ,TD036 AS [TF132]
                                            ,TD037 AS [TF133]
                                            ,TD038 AS [TF134]
                                            ,TD014 AS [TF135]
                                            ,0 AS [TF136]
                                            ,0 AS [TF137]
                                            ,'' AS [TF138]
                                            ,'' AS [TF139]
                                            ,'' AS [TF140]
                                            ,0 AS [TF141]
                                            ,'' AS [TF142]
                                            ,'' AS [TF143]
                                            ,'' AS [TF144]
                                            ,'2' AS [TF145]
                                            ,'2' AS [TF146]
                                            ,'' AS [TF147]
                                            ,'' AS [TF148]
                                            ,'' AS [TF149]
                                            ,'' AS [TF150]
                                            ,'' AS [TF151]
                                            ,TD080 AS [TF152]
                                            ,TD081 AS [TF153]
                                            ,TD082 AS [TF154]
                                            ,TD083 AS [TF155]
                                            ,TD080 AS [TF156]
                                            ,TD081 AS [TF157]
                                            ,TD082 AS [TF158]
                                            ,TD083 AS [TF159]
                                            ,TD084 AS [TF160]
                                            ,TD085 AS [TF161]
                                            ,TD084 AS [TF162]
                                            ,TD085 AS [TF163]
                                            ,0 AS [TF164]
                                            ,0 AS [TF165]
                                            ,0 AS [TF166]
                                            ,0 AS [TF167]
                                            ,'' AS [TF168]
                                            ,'' AS [TF169]
                                            ,'' AS [TF170]
                                            ,'' AS [TF171]
                                            ,'' AS [TF172]
                                            ,'' AS [TF173]
                                            ,CONVERT(NVARCHAR,[PURTATBCHAGE].VERSIONS)+CONVERT(NVARCHAR,[PURTATBCHAGE].TA001)+CONVERT(NVARCHAR,[PURTATBCHAGE].TA002)+CONVERT(NVARCHAR,[PURTATBCHAGE].TB003) AS [UDF01]
                                            ,'' AS [UDF02]
                                            ,'' AS [UDF03]
                                            ,'' AS [UDF04]
                                            ,'' AS [UDF05]
                                            ,0 AS [UDF06]
                                            ,0 AS [UDF07]
                                            ,0 AS [UDF08]
                                            ,0 AS [UDF09]
                                            ,0 AS [UDF10]
                                            FROM [TK].dbo.PURTD,[TKPUR].[dbo].[PURTATBCHAGE]
                                            WHERE 1=1
                                            AND PURTD.TD026=[PURTATBCHAGE].TA001 AND PURTD.TD027=[PURTATBCHAGE].TA002 AND PURTD.TD028=[PURTATBCHAGE].TB003
                                            AND TD001='{4}' AND TD002='{5}'
                                            AND [PURTATBCHAGE].TA001='{0}' AND [PURTATBCHAGE].TA002='{1}' AND [PURTATBCHAGE].VERSIONS='{2}'
                                            INSERT INTO [TK].[dbo].[PURTE]
                                            (
                                            [COMPANY]
                                            ,[CREATOR]
                                            ,[USR_GROUP]
                                            ,[CREATE_DATE]
                                            ,[MODIFIER]
                                            ,[MODI_DATE]
                                            ,[FLAG]
                                            ,[CREATE_TIME]
                                            ,[MODI_TIME]
                                            ,[TRANS_TYPE]
                                            ,[TRANS_NAME]
                                            ,[sync_date]
                                            ,[sync_time]
                                            ,[sync_mark]
                                            ,[sync_count]
                                            ,[DataUser]
                                            ,[DataGroup]
                                            ,[TE001]
                                            ,[TE002]
                                            ,[TE003]
                                            ,[TE004]
                                            ,[TE005]
                                            ,[TE006]
                                            ,[TE007]
                                            ,[TE008]
                                            ,[TE009]
                                            ,[TE010]
                                            ,[TE011]
                                            ,[TE012]
                                            ,[TE013]
                                            ,[TE014]
                                            ,[TE015]
                                            ,[TE016]
                                            ,[TE017]
                                            ,[TE018]
                                            ,[TE019]
                                            ,[TE020]
                                            ,[TE021]
                                            ,[TE022]
                                            ,[TE023]
                                            ,[TE024]
                                            ,[TE025]
                                            ,[TE026]
                                            ,[TE027]
                                            ,[TE028]
                                            ,[TE029]
                                            ,[TE030]
                                            ,[TE031]
                                            ,[TE032]
                                            ,[TE033]
                                            ,[TE034]
                                            ,[TE035]
                                            ,[TE036]
                                            ,[TE037]
                                            ,[TE038]
                                            ,[TE039]
                                            ,[TE040]
                                            ,[TE041]
                                            ,[TE042]
                                            ,[TE043]
                                            ,[TE045]
                                            ,[TE046]
                                            ,[TE047]
                                            ,[TE048]
                                            ,[TE103]
                                            ,[TE107]
                                            ,[TE108]
                                            ,[TE109]
                                            ,[TE110]
                                            ,[TE113]
                                            ,[TE114]
                                            ,[TE115]
                                            ,[TE118]
                                            ,[TE119]
                                            ,[TE120]
                                            ,[TE121]
                                            ,[TE122]
                                            ,[TE123]
                                            ,[TE124]
                                            ,[TE125]
                                            ,[TE134]
                                            ,[TE135]
                                            ,[TE136]
                                            ,[TE137]
                                            ,[TE138]
                                            ,[TE139]
                                            ,[TE140]
                                            ,[TE141]
                                            ,[TE142]
                                            ,[TE143]
                                            ,[TE144]
                                            ,[TE145]
                                            ,[TE146]
                                            ,[TE147]
                                            ,[TE148]
                                            ,[TE149]
                                            ,[TE150]
                                            ,[TE151]
                                            ,[TE152]
                                            ,[TE153]
                                            ,[TE154]
                                            ,[TE155]
                                            ,[TE156]
                                            ,[TE157]
                                            ,[TE158]
                                            ,[TE159]
                                            ,[TE160]
                                            ,[TE161]
                                            ,[TE162]
                                            ,[UDF01]
                                            ,[UDF02]
                                            ,[UDF03]
                                            ,[UDF04]
                                            ,[UDF05]
                                            ,[UDF06]
                                            ,[UDF07]
                                            ,[UDF08]
                                            ,[UDF09]
                                            ,[UDF10]
                                            )
                                            SELECT 
                                            PURTC.[COMPANY]
                                            ,PURTC.[CREATOR]
                                            ,PURTC.[USR_GROUP]
                                            ,PURTC.[CREATE_DATE]
                                            ,PURTC.[MODIFIER]
                                            ,PURTC.[MODI_DATE]
                                            ,PURTC.[FLAG]
                                            ,PURTC.[CREATE_TIME]
                                            ,PURTC.[MODI_TIME]
                                            ,PURTC.[TRANS_TYPE]
                                            ,'PURI08' AS [TRANS_NAME]
                                            ,PURTC.[sync_date]
                                            ,PURTC.[sync_time]
                                            ,PURTC.[sync_mark]
                                            ,PURTC.[sync_count]
                                            ,PURTC.[DataUser]
                                            ,PURTC.[DataGroup]
                                            ,TC001 AS [TE001]
                                            ,TC002 AS [TE002]
                                            ,'{3}' AS [TE003]
                                            ,CONVERT(NVARCHAR,GETDATE(),112) AS [TE004]
                                            ,TC004 AS [TE005]
                                            ,'' AS [TE006]
                                            ,TC005 AS [TE007]
                                            ,TC006 AS [TE008]
                                            ,TC007 AS [TE009]
                                            ,TC008 AS [TE010]
                                            ,CONVERT(NVARCHAR,GETDATE(),112) AS [TE011]
                                            ,'N' AS [TE012]
                                            ,TC015 AS [TE013]
                                            ,TC016 AS [TE014]
                                            ,TC017 AS [TE015]
                                            ,0 AS [TE016]
                                            ,'N' AS [TE017]
                                            ,TC018 AS [TE018]
                                            ,TC021 AS [TE019]
                                            ,TC022 AS [TE020]
                                            ,'' AS [TE021]
                                            ,TC026 AS [TE022]
                                            ,TC027 AS [TE023]
                                            ,TC028 AS [TE024]
                                            ,'N' AS [TE025]
                                            ,0 AS [TE026]
                                            ,TC009 AS [TE027]
                                            ,'N' AS [TE028]
                                            ,TC035 AS [TE029]
                                            ,'' AS [TE030]
                                            ,'' AS [TE031]
                                            ,'N' AS [TE032]
                                            ,'' AS [TE033]
                                            ,0 AS [TE034]
                                            ,0 AS [TE035]
                                            ,'' AS [TE036]
                                            ,TC011 AS [TE037]
                                            ,'' AS [TE038]
                                            ,'' AS [TE039]
                                            ,'' AS [TE040]
                                            ,TC050 AS [TE041]
                                            ,'' AS [TE042]
                                            ,TC036 AS [TE043]
                                            ,TC037 AS [TE045]
                                            ,TC038 AS [TE046]
                                            ,TC039 AS [TE047]
                                            ,TC040 AS [TE048]
                                            ,'' AS [TE103]
                                            ,TC005 AS [TE107]
                                            ,TC006 AS [TE108]
                                            ,TC007 AS [TE109]
                                            ,TC008 AS [TE110]
                                            ,TC015 AS [TE113]
                                            ,TC016 AS [TE114]
                                            ,TC017 AS [TE115]
                                            ,TC018 AS [TE118]
                                            ,TC021 AS [TE119]
                                            ,TC022 AS [TE120]
                                            ,TC026 AS [TE121]
                                            ,TC027 AS [TE122]
                                            ,TC028 AS [TE123]
                                            ,TC009 AS [TE124]
                                            ,TC035 AS [TE125]
                                            ,0 AS [TE134]
                                            ,0 AS [TE135]
                                            ,'' AS [TE136]
                                            ,'' AS [TE137]
                                            ,'' AS [TE138]
                                            ,'' AS [TE139]
                                            ,'1' AS [TE140]
                                            ,'N' AS [TE141]
                                            ,'N' AS [TE142]
                                            ,TC036 AS [TE143]
                                            ,'N' AS [TE144]
                                            ,'' AS [TE145]
                                            ,TC041 AS [TE146]
                                            ,TC041 AS [TE147]
                                            ,TC011 AS [TE148]
                                            ,0 AS [TE149]
                                            ,0 AS [TE150]
                                            ,0 AS [TE151]
                                            ,0 AS [TE152]
                                            ,'' AS [TE153]
                                            ,'' AS [TE154]
                                            ,'' AS [TE155]
                                            ,'' AS [TE156]
                                            ,'' AS [TE157]
                                            ,'' AS [TE158]
                                            ,TC037 AS [TE159]
                                            ,TC038 AS [TE160]
                                            ,TC039 AS [TE161]
                                            ,TC040 AS [TE162]
                                            ,'Y' AS [UDF01]
                                            ,'' AS [UDF02]
                                            ,'' AS [UDF03]
                                            ,'' AS [UDF04]
                                            ,'' AS [UDF05]
                                            ,0 AS [UDF06]
                                            ,0 AS [UDF07]
                                            ,0 AS [UDF08]
                                            ,0 AS [UDF09]
                                            ,0 AS [UDF10]
                                            FROM  [TK].dbo.PURTC
                                            WHERE 1=1
                                            AND TC001='{4}' AND TC002='{5}'
                                            ", DR["TA001"].ToString(), DR["TA002"].ToString(), DR["VERSIONS"].ToString(), DR["TE003"].ToString(), DR["TE001"].ToString(), DR["TE002"].ToString());
                    }

                    sbSql.AppendFormat(@"  
                                   
                                        ");

                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;


                    cmd.CommandText = sbSql.ToString();
                    cmd.Transaction = tran;
                    int result = cmd.ExecuteNonQuery();

                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消
                    }
                    else
                    {
                        tran.Commit();      //執行交易  


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

        public void SEARCHPURTE(string TA001, string TA002, string VERSIONS)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();

            try
            {
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ConnectionString);
                SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString);
                StringBuilder sbSql = new StringBuilder();
                StringBuilder sbSqlQuery = new StringBuilder();


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                 
                                    SELECT TE001 AS '採購變更單別',TE002 AS '採購變更單號',TE003 AS '版次'
                                    FROM [TK].dbo.PURTE
                                    WHERE TE001+TE002 IN 
                                    (
                                    SELECT TD001+TD002
                                    FROM [TK].dbo.PURTD
                                    WHERE TD026+TD027+TD028 IN 
                                    (
                                    SELECT TA001+TA002+TB003
                                    FROM [TKPUR].[dbo].[PURTATBCHAGE]
                                    WHERE  TA001='{0}' AND TA002='{1}' AND VERSIONS='{2}'
                                    )
                                    GROUP BY  TD001,TD002
                                    )
                                    ", TA001, TA002, VERSIONS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    //dataGridView6.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        //dataGridView6.DataSource = ds.Tables["ds"];
                        //dataGridView6.AutoResizeColumns();

                    }

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
