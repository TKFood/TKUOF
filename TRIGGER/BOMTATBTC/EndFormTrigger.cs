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

namespace TKUOF.TRIGGER.BOMTATBTC
{
    //BOM10.BOM變更單 的核準

    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {
            
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TA001=null;
            string TA002 = null;
          
            string FORMID= null;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TA001 = applyTask.Task.CurrentDocument.Fields["TA001"].FieldValue.ToString().Trim();
            TA002 = applyTask.Task.CurrentDocument.Fields["TA002"].FieldValue.ToString().Trim();
           
            FORMID = applyTask.FormNumber;

            ///核準
            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                if (!string.IsNullOrEmpty(TA001) && !string.IsNullOrEmpty(TA002) )
                {
                    UPDATEBOMMCBOMMD(TA001, TA002, FORMID);
                }
            }
            //作廢
            else if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Cancel)
            {
                //if (!string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002))
                //{
                //    UPDATEPURTCPURTDCANCEL(TC001, TC002, FORMID);
                //}
            }
            //否決
            else if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Reject)
            {
            //    if (!string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002))
            //    {
            //        UPDATEPURTCPURTDREJECT(TC001, TC002, FORMID);
            //    }
            }
            //退簽
            else if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Return)
            {
                //if (!string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002))
                //{
                //    UPDATEPURTCPURTDRETURN(TC001, TC002, FORMID);
                //}
            }

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public void UPDATEBOMMCBOMMD(string TA001, string TA002, string FORMID)
        {
            DataTable dt = SEARCHBOMTABOMTBBOMTC(TA001, TA002);

            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");


            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();
            
            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                        ");

            foreach (DataRow DRDATA in dt.Rows)
            {
                queryString.AppendFormat(@"

                                            UPDATE [test0923].dbo.BOMMC
                                            SET MC001=TB004
                                            ,MC002=TB005
                                            ,MC003=TB006
                                            ,MC004=TB008
                                            ,MC005=TB009
                                            ,MC006=TB001
                                            ,MC007=TB002
                                            ,MC008=TB003
                                            ,MC009=TB007
                                            ,MC010=TB010
                                            ,MODI_DATE=TA003
                                            ,FLAG=FLAG+1 
                                            ,COMPANY=TEMP.COMPANY
                                            ,MODIFIER=TEMP.MODIFIER 
                                            ,MODI_TIME=TEMP.MODI_TIME 
                                            ,MC024=MB002 
                                            ,MC025=MB003
                                            FROM 
                                            (
                                            SELECT BOMTA.COMPANY,BOMTA.MODIFIER,BOMTA.MODI_TIME,TA001,TA002,TA003,TA005,TA006
                                            ,TB001,TB002,TB003,TB004,TB005,TB006,TB007,TB008,TB009,TB010,TB104
                                            ,MB002,MB003
                                            FROM [test0923].dbo.BOMTA,[test0923].dbo.BOMTB,[test0923].dbo.INVMB
                                            WHERE TA001=TB001 AND TA002=TB002 
                                            AND TB004=MB001
                                            AND TA001='{0}' AND TA002='{1}'
                                            AND TB004='{2}'
                                            )  AS TEMP 
                                            WHERE MC001=TEMP.TB004
                                            AND MC001='{2}'

                                           
                                            INSERT INTO [test0923].dbo.BOMMD 
                                            (MD001,MD002,MD003,MD004,MD005,MD006,MD007,MD008,MD009,MD010 
                                            ,MD011,MD012,MD013,MD014,MD015,MD016,MD017,MD018,MD019,MD020
                                            ,MD021,MD022,MD023,MD029,MD035,MD036
                                            ,COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME ) 

                                            SELECT 
                                            TB004 MD001,TC004 MD002,TC005 MD003,TC006 MD004,TC007 MD005,TC008 MD006,TC009 MD007,TC010 MD008,TC011 MD009,TC012 MD010 
                                            ,TC013 MD011,TC014 MD012,TC015 MD013,TC016 MD014,TC017 MD015,TC018 MD016,TC019 MD017,TC020 MD018,TC021 MD019,TC022 MD020
                                            ,TC023 MD021,TC024 MD022,TC025 MD023,TC032 MD029,MB002 MD035,MB003 MD036
                                            ,BOMTA.COMPANY ,BOMTA.CREATOR ,BOMTA.USR_GROUP ,BOMTA.CREATE_DATE ,BOMTA.FLAG, BOMTA.CREATE_TIME, BOMTA.MODI_TIME, 'P004' TRANS_TYPE, 'BOMI04' TRANS_NAME 

                                            FROM [test0923].dbo.BOMTA,[test0923].dbo.BOMTB,[test0923].dbo.BOMTC
                                            LEFT JOIN [test0923].dbo.INVMB ON TC005=MB001
                                            WHERE TA001=TB001 AND TA002=TB002 AND TA001=TC001 AND TA002=TC002 AND TB003=TC003
                                            AND TA001='{0}' AND TA002='{1}'
                                            AND TB004='{2}'
                                            AND TC004 NOT IN (SELECT MD002 FROM [test0923].dbo.BOMMD WHERE  MD001='{2}')


                                        
                                            UPDATE [test0923].dbo.BOMMD  
                                            SET MD001=TB004 
                                            ,MD002=TC004 
                                            ,MD003=TC005
                                            ,MD004=TC006
                                            ,MD005=TC007
                                            ,MD006=TC008 
                                            ,MD007=TC009 
                                            ,MD008=TC010
                                            ,MD009=TC011
                                            ,MD010=TC012
                                            ,MD011=TC013
                                            ,MD012=TC014
                                            ,MD013=TC015
                                            ,MD014=TC016
                                            ,MD015=TC017
                                            ,MD016=TC018
                                            ,MD017=TC019
                                            ,MD018=TC020
                                            ,MD019=TC021
                                            ,MD020=TC022
                                            ,MD021=TC023
                                            ,MD022=TC024
                                            ,MD023=TC025
                                            ,MD029=TC032
                                            ,MD035=MB002
                                            ,MD036=MB003  
                                            ,FLAG=FLAG+1
                                            ,COMPANY=TEMP.COMPANY
                                            ,MODIFIER=TEMP.MODIFIER
                                            ,MODI_DATE=TEMP.MODI_DATE  
                                            ,MODI_TIME=TEMP.MODI_TIME 
                                            FROM 
                                            (
                                            SELECT BOMTA.COMPANY,BOMTA.MODIFIER,BOMTA.MODI_DATE,BOMTA.MODI_TIME,TA001,TA002,TA003,TA005,TA006
                                            ,TB004,TB005,TB006,TB007,TB008,TB009,TB010,TB104
                                            ,TC004,TC005,TC006,TC007,TC008,TC009,TC010
                                            ,TC011,TC012,TC013,TC014,TC015,TC016,TC017,TC018,TC019,TC020
                                            ,TC021,TC022,TC023,TC024,TC025,TC032
                                            ,TC104,TC105
                                            ,MB002,MB003
                                            FROM [test0923].dbo.BOMMD,[test0923].dbo.BOMTA,[test0923].dbo.BOMTB,[test0923].dbo.BOMTC
                                            LEFT JOIN [test0923].dbo.INVMB ON TC005=MB001
                                            WHERE TA001=TB001 AND TA002=TB002 AND TA001=TC001 AND TA002=TC002 AND TB003=TC003
                                            AND MD001=TB004 AND MD002=TC004
                                            AND ISNULL(TC005,'')<>''
                                            AND TA001='{0}' AND TA002='{1}'
                                            AND TB004='{2}'
                                            ) AS TEMP
                                            WHERE MD001=TEMP.TB004
                                            AND MD002=TEMP.TC004
                                            AND MD001='{2}'

                                       
                                            DELETE [test0923].dbo.BOMMD
                                            FROM (
                                            SELECT TB004,TC004
                                            FROM [test0923].dbo.BOMMD,[test0923].dbo.BOMTA,[test0923].dbo.BOMTB,[test0923].dbo.BOMTC
                                            LEFT JOIN [test0923].dbo.INVMB ON TC005=MB001
                                            WHERE TA001=TB001 AND TA002=TB002 AND TA001=TC001 AND TA002=TC002 AND TB003=TC003
                                            AND MD001=TB004 AND MD002=TC004
                                            AND ISNULL(TC005,'')=''
                                            AND TA001='{0}' AND TA002='{1}'
                                            AND TB004='{2}'
                                            ) AS TEMP
                                            WHERE BOMMD.MD001=TEMP.TB004
                                            AND BOMMD.MD002=TEMP.TC004
                                            AND BOMMD.MD001='{2}'


                                        ", DRDATA["TA001"].ToString(), DRDATA["TA002"].ToString(), DRDATA["TB004"].ToString());
            }

           
            

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);

                    //command.Parameters.Add("@TA001", SqlDbType.NVarChar).Value = TA001;
                    //command.Parameters.Add("@TA002", SqlDbType.NVarChar).Value = TA002;                   


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

        public DataTable SEARCHBOMTABOMTBBOMTC(string TA001, string TA002)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();
            Ede.Uof.Utility.Data.DatabaseHelper m_db = new Ede.Uof.Utility.Data.DatabaseHelper(connectionString);

            string cmdTxt = @" 
                            SELECT TA001,TA002,TA003,TA005,TA006
                            ,TB001,TB002,TB003,TB004,TB005,TB006,TB007,TB008,TB009,TB010,TB104
                            ,MB002,MB003
                            FROM [test0923].dbo.BOMTA,[test0923].dbo.BOMTB,[test0923].dbo.INVMB
                            WHERE TA001=TB001 AND TA002=TB002 
                            AND TB004=MB001
                            AND TA001=@TA001 AND TA002=@TA002
                        ";

            m_db.AddParameter("@TA001", TA001);
            m_db.AddParameter("@TA002", TA002);

            DataTable dt = new DataTable();

            dt.Load(m_db.ExecuteReader(cmdTxt));

            if (dt.Rows.Count > 0)
            {
                return dt;
            }
            else
            {
                return null;
            }
           
        }
    }
}
