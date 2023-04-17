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
using Ede.Uof.EIP.Organization.Util;
using Ede.Uof.EIP.SystemInfo;

namespace TKUOF.TRIGGER.INVI05
{
    //INVTA INVTB 庫存異動單的 核準


    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {

        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TA001 = null;
            string TA002 = null;
            string TA003 = null;
            string FORMID = null;
            string MODIFIER = null;
            UserUCO userUCO = new UserUCO();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TA001 = applyTask.Task.CurrentDocument.Fields["TA001"].FieldValue.ToString().Trim();
            TA002 = applyTask.Task.CurrentDocument.Fields["TA002"].FieldValue.ToString().Trim();
            TA003 = applyTask.Task.CurrentDocument.Fields["TA003"].FieldValue.ToString().Trim();

            FORMID = applyTask.FormNumber;
            //MODIFIER = applyTask.Task.Applicant.Account;

            //取得簽核人工號
            EBUser ebUser = userUCO.GetEBUser(Current.UserGUID);
            MODIFIER = ebUser.Account;

           
            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(TA001) && !string.IsNullOrEmpty(TA002) )
                {
                    UPDATE_INVI05(TA001, TA002, TA003,FORMID, MODIFIER);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATE_INVI05(string TA001, string TA002, string TA003, string FORMID, string MODIFIER)
        {           
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            string TA006 = "Y";
            string TA015 = MODIFIER;
            string TA017 = "N";
            string TB018 = "Y";
            string TB019 = TA003;
            string MC013 = TA003;


            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"   

                                        UPDATE [test0923].dbo.INVTA
                                        SET
                                        TA006=@TA006
                                        ,TA015=@MODIFIER
                                        ,TA017=@TA017
                                        ,FLAG=FLAG+1
                                        ,MODIFIER=@MODIFIER
                                        ,MODI_DATE=@MODI_DATE
                                        ,MODI_TIME=@MODI_TIME
                                        WHERE  TA001=@TA001 and TA002=@TA002                                   

                                        UPDATE [test0923].dbo.INVTB
                                        SET  
                                        TB018=@TB018 
                                        ,TB019=@TB019 
                                        ,FLAG=FLAG+1
                                        ,MODIFIER=@MODIFIER
                                        ,MODI_DATE=@MODI_DATE
                                        ,MODI_TIME=@MODI_TIME 
                                        WHERE TB001=@TB001 AND TB002=@TB002

                                        UPDATE [test0923].dbo.INVMB
                                        SET 
                                        MB064=MB064+TEMP.TB007
                                        ,MB065=MB065+TEMP.TB007MONEY
                                        ,MB089=MB089+TEMP.TB022
                                        ,INVMB.FLAG=INVMB.FLAG+1
                                        FROM 
                                        (
                                        SELECT TB004,SUM(TB007) TB007,SUM(TB022) TB022,(CASE WHEN MB064>0 AND MB065>0 THEN SUM((MB065/MB064)*TB007) ELSE 0 END) AS TB007MONEY
                                        FROM [test0923].dbo.INVTA,[test0923].dbo.INVTB,[test0923].dbo.INVMB
                                        WHERE 1=1
                                        AND TB004=MB001
                                        AND TA001=TB001 AND TA002=TB002
                                        AND TA001=@TA001 AND TA002=@TA002
                                        GROUP BY TB004,MB064,MB065
                                        ) AS TEMP
                                        WHERE TEMP.TB004=INVMB.MB001

                                        UPDATE [test0923].dbo.INVMC
                                        SET 
                                        MC007=MC007+TEMP.TB007
                                        ,MC008=MC008+TEMP.TB007MONEY
                                        ,MC013=@MC013 
                                        ,MC014=MC014+TEMP.TB022
                                        ,INVMC.FLAG=INVMC.FLAG+1
                                        FROM 
                                        (
                                        SELECT TB004,TB012,SUM(TB007) TB007,SUM(TB022) TB022,(CASE WHEN MB064>0 AND MB065>0 THEN SUM((MB065/MB064)*TB007) ELSE 0 END) AS TB007MONEY
                                        FROM [test0923].dbo.INVTA,[test0923].dbo.INVTB,[test0923].dbo.INVMB
                                        WHERE 1=1
                                        AND TB004=MB001
                                        AND TA001=TB001 AND TA002=TB002
                                        AND TA001=@TA001 AND TA002=@TA002
                                        GROUP BY TB004,TB012,MB064,MB065

                                        ) AS TEMP
                                        WHERE TEMP.TB004=MC001 AND TEMP.TB012=MC002

                                        INSERT INTO [test0923].dbo.INVLA
                                        (LA001 ,LA002 , LA003 ,LA004 ,LA005 ,LA006,LA007,LA008 ,LA009 ,LA010 , 
                                        LA011 ,LA012 ,LA013 ,LA014 ,LA015 ,LA016 ,LA017,LA018,LA019,LA020,LA021, 
                                        COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
                                        CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME ) 
                                        SELECT 
                                        TB004 LA001 ,'' LA002 ,'' LA003 ,TA003 LA004 , MQ010 LA005 ,TB001 LA006,TB002 LA007,TB003 LA008 ,TB012 LA009 ,TB017 LA010 , 
                                        TB007 LA011 ,(CASE WHEN MB065>0 AND MB064>0 THEN MB065/MB064 ELSE 0 END ) LA012 ,(CASE WHEN MB065>0 AND MB064>0 THEN MB065/MB064*TB007 ELSE 0 END ) LA013 ,MQ008 LA014 ,MQ011 LA015 ,TB014 LA016 ,(CASE WHEN MB065>0 AND MB064>0 THEN MB065/MB064*TB007 ELSE 0 END ) LA017,0 LA018,0 LA019,0 LA020,0 LA021, 
                                        INVTA.COMPANY COMPANY ,INVTA.CREATOR CREATOR ,INVTA.USR_GROUP USR_GROUP ,INVTA.CREATE_DATE CREATE_DATE ,0 FLAG, 
                                        INVTA.CREATE_TIME CREATE_TIME,INVTA.MODI_TIME MODI_TIME,INVTA.TRANS_TYPE TRANS_TYPE,INVTA.TRANS_NAME TRANS_NAME 
                                        FROM [test0923].dbo.INVTA,[test0923].dbo.INVTB,[test0923].dbo.INVMB,[test0923].dbo.CMSMQ
                                        WHERE 1=1
                                        AND TB004=MB001
                                        AND TA001=MQ001
                                        AND TA001=TB001 AND TA002=TB002
                                        AND TA001=@TA001 AND TA002=@TA002

                                        INSERT INTO [test0923].dbo.INVMF 
                                        (MF001 ,MF002 ,MF003 ,MF004 ,MF005 ,MF006,MF007,MF008 ,MF009 ,MF010 , 
                                        MF011 ,MF012 ,MF013,MF014 ,COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
                                        CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME ) 
                                        SELECT 
                                        TB004 MF001 ,TB014 MF002 ,TA003 MF003 ,TB001 MF004 ,TB002 MF005 ,TB003 MF006,TB012 MQ010,MQ010 MF008 ,MQ008 MF009 ,TB007 MF010 , 
                                        '' MF011 ,'' MF012 ,TB017 MF013,TB022 MF014 
                                        ,INVTA.COMPANY COMPANY ,INVTA.CREATOR CREATOR ,INVTA.USR_GROUP USR_GROUP ,INVTA.CREATE_DATE CREATE_DATE ,0 FLAG, 
                                        INVTA.CREATE_TIME CREATE_TIME,INVTA.MODI_TIME MODI_TIME,INVTA.TRANS_TYPE TRANS_TYPE,INVTA.TRANS_NAME TRANS_NAME
                                        FROM [test0923].dbo.INVTA,[test0923].dbo.INVTB,[test0923].dbo.INVMB,[test0923].dbo.CMSMQ
                                        WHERE 1=1
                                        AND TB004=MB001
                                        AND TA001=MQ001
                                        AND TA001=TB001 AND TA002=TB002
                                        AND TA001=@TA001 AND TA002=@TA002



                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TA001", SqlDbType.NVarChar).Value = TA001;
                    command.Parameters.Add("@TA002", SqlDbType.NVarChar).Value = TA002;

                    command.Parameters.Add("@TB001", SqlDbType.NVarChar).Value = TA001;
                    command.Parameters.Add("@TB002", SqlDbType.NVarChar).Value = TA002;

                    command.Parameters.Add("@TA006", SqlDbType.NVarChar).Value = TA006;
                    command.Parameters.Add("@TA015", SqlDbType.NVarChar).Value = TA015;
                    command.Parameters.Add("@TA017", SqlDbType.NVarChar).Value = TA017;
                    command.Parameters.Add("@TB018", SqlDbType.NVarChar).Value = TB018;
                    command.Parameters.Add("@TB019", SqlDbType.NVarChar).Value = TB019;
                    command.Parameters.Add("@MC013", SqlDbType.NVarChar).Value = MC013;


                    command.Parameters.Add("@MODIFIER", SqlDbType.NVarChar).Value = MODIFIER;
                    command.Parameters.Add("@MODI_DATE", SqlDbType.NVarChar).Value = DateTime.Now.ToString("yyyyMMdd");
                    command.Parameters.Add("@MODI_TIME", SqlDbType.NVarChar).Value = DateTime.Now.ToString("HH:mm:ss");
                    command.Parameters.Add("@UDF02", SqlDbType.NVarChar).Value = FORMID;



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
