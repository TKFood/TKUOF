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

namespace TKUOF.TRIGGER.COP40
{
    //核準


    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {

        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TI001 = null;
            string TI002 = null;
            string FORMID = null;
            string MODIFIER = null;
            UserUCO userUCO = new UserUCO();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TI001 = applyTask.Task.CurrentDocument.Fields["TI001"].FieldValue.ToString().Trim();
            TI002 = applyTask.Task.CurrentDocument.Fields["TI002"].FieldValue.ToString().Trim();

            FORMID = applyTask.FormNumber;
            //MODIFIER = applyTask.Task.Applicant.Account;

            //取得簽核人工號
            EBUser ebUser = userUCO.GetEBUser(Current.UserGUID);
            MODIFIER = ebUser.Account;

           
            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(TI001) && !string.IsNullOrEmpty(TI002) )
                {
                    UPDATE_COP40(TI001, TI002, FORMID, MODIFIER);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATE_COP40(string TI001, string TI002, string FORMID, string MODIFIER)
        {           
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            string TI019 = "Y";
            string TI041 = "N";
            string TI035 = MODIFIER;
            string TJ021 = "Y";
            string TJ024 = "Y";

            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"       
UPDATE [test0923].dbo.COPTI 
SET 
TI019=@TI019 
,TI035=@TI035
,TI041=@TI041 
,FLAG=FLAG+1
,MODIFIER=@MODIFIER
,MODI_DATE=@MODI_DATE
,MODI_TIME=@MODI_TIME 
WHERE TI001=@TI001 AND TI002=@TI002 

UPDATE [test0923].dbo.COPTJ 
SET 
TJ021=@TJ021
,TJ024=@TJ024
,FLAG=FLAG+1
,MODIFIER=@MODIFIER
,MODI_DATE=@MODI_DATE
,MODI_TIME=@MODI_TIME 
WHERE TJ001=@TJ001 AND TJ002=@TJ002


INSERT INTO [test0923].dbo.INVLA 
(LA001 ,LA002 , LA003 ,LA004 ,LA005 ,LA006,LA007,LA008 ,LA009 ,LA010 , 
LA011 ,LA012 ,LA013 ,LA014 ,LA015 ,LA016 ,LA017,LA018,LA019,LA020,LA021, 
COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME ) 
SELECT 
TJ001 LA001 ,'' LA002 ,'' LA003 ,TI003 LA004 ,1 LA005 ,TJ001 LA006, TJ002 LA007,TJ003 LA008 ,TJ013 LA009 ,TJ023 LA010 , 
(CASE WHEN TJ030='1' THEN (TJ007+TJ047) ELSE '0' END) LA011 ,(CASE WHEN MB064>0 AND MB065>0 THEN MB065/MB064 ELSE 0 END) LA012 ,(CASE WHEN MB064>0 AND MB065>0 THEN MB065/MB064 ELSE 0 END)* (CASE WHEN TJ030='1' THEN (TJ007+TJ047) ELSE '0' END) LA013 ,'2' LA014 ,'N' LA015 ,TJ014 LA016 ,(CASE WHEN MB064>0 AND MB065>0 THEN MB065/MB064 ELSE 0 END)* (TJ007+TJ047)  LA017,0 LA018,0 LA019,0 LA020,0 LA021, 
COPTI.COMPANY ,COPTI.CREATOR ,COPTI.USR_GROUP ,COPTI.CREATE_DATE ,0 FLAG, 
COPTI.CREATE_TIME, COPTI.MODI_TIME, COPTI.TRANS_TYPE, COPTI.TRANS_NAME
FROM [test0923].dbo.COPTI, [test0923].dbo.COPTJ,[test0923].dbo.INVMB
WHERE TI001=TJ001 AND TI002=TJ002
AND TJ004=MB001
AND TI001=@TI001 AND TI002=@TI002

INSERT INTO [test0923].dbo.INVMF 
(MF001 ,MF002 ,MF003 ,MF004 ,MF005 ,MF006,MF007,MF008 ,MF009 ,MF010 , 
MF011 ,MF012 ,MF013,MF014 ,COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME ) 
SELECT 
TJ004 MF001 ,TJ014 MF002 ,TI003 MF003 ,TJ001 MF004 ,TJ002 MF005 ,TJ003 MF006,TJ013 MF007,'1' MF008 ,'2' MF009 ,(CASE WHEN TJ030='1' THEN (TJ007+TJ047) ELSE '0' END) MF010 , 
'' MF011 ,'' MF012 ,''  MF013,0 MF014 
,COPTI.COMPANY ,COPTI.CREATOR ,COPTI.USR_GROUP ,COPTI.CREATE_DATE ,0 FLAG, 
COPTI.CREATE_TIME, COPTI.MODI_TIME, COPTI.TRANS_TYPE, COPTI.TRANS_NAME
FROM [test0923].dbo.COPTI, [test0923].dbo.COPTJ,[test0923].dbo.INVMB
WHERE TI001=TJ001 AND TI002=TJ002
AND TJ004=MB001
AND TI001=@TI001 AND TI002=@TI002

UPDATE [test0923].dbo.INVMB 
SET INVMB.MB064=INVMB.MB064+TEMP.MB064, INVMB.MB065=INVMB.MB065+TEMP.MB065,INVMB.MB089=INVMB.MB089+TEMP.MB089,INVMB.FLAG=INVMB.FLAG+1
FROM (
SELECT 
MB001 
,(CASE WHEN TJ030='1' THEN SUM(TJ007+TJ047) ELSE '0' END) MB064 
,((CASE WHEN TJ030='1' THEN SUM(TJ007+TJ047) ELSE '0' END)*(CASE WHEN MB064>0 AND MB065>0 THEN MB065/MB064 ELSE 0 END)) MB065
,(CASE WHEN TJ030='1' THEN SUM(TJ035+TJ048) ELSE '0' END)MB089
FROM [test0923].dbo.COPTI, [test0923].dbo.COPTJ,[test0923].dbo.INVMB
WHERE TI001=TJ001 AND TI002=TJ002
AND TJ004=MB001
AND TI001=@TI001 AND TI002=@TI002
GROUP BY MB001,TJ030,MB064,MB065
) AS TEMP 
WHERE TEMP.MB001=INVMB.MB001

UPDATE [test0923].dbo.INVMC 
SET INVMC.MC008=INVMC.MC008+TEMP.MC008,INVMC.MC007=INVMC.MC007++TEMP.MC007,INVMC.MC012=TI003,INVMC.MC014=INVMC.MC014+TEMP.MC014
,INVMC.FLAG=INVMC.FLAG+1
FROM (
SELECT 
MB001,TJ013 ,TI003
,(CASE WHEN TJ030='1' THEN SUM(TJ007+TJ047) ELSE '0' END) MC007 
,((CASE WHEN TJ030='1' THEN SUM(TJ007+TJ047) ELSE '0' END)*(CASE WHEN MB064>0 AND MB065>0 THEN MB065/MB064 ELSE 0 END)) MC008
,(CASE WHEN TJ030='1' THEN SUM(TJ035+TJ048) ELSE '0' END) MC014
FROM [test0923].dbo.COPTI, [test0923].dbo.COPTJ,[test0923].dbo.INVMB
WHERE TI001=TJ001 AND TI002=TJ002
AND TJ004=MB001
AND TI001=@TI001 AND TI002=@TI002
GROUP BY MB001,TJ030,MB064,MB065,TJ013,TI003
) AS TEMP 
WHERE TEMP.MB001=INVMC.MC001 AND TEMP.TJ013=INVMC.MC002

UPDATE [test0923].dbo.COPTD 
SET TD009=TD009-NUMSTJ007,TD016='N',FLAG=FLAG+1 ,TD033=TD033-NUMSTJ035,TD080='1'
,TD025=TD025-NUMSTJ047,TD035=TD035-NUMSTJ048 
FROM (
SELECT 
TJ018,TJ019,TJ020,MB001,TJ013 ,TI003
,(CASE WHEN TJ030='1' THEN (TJ007) ELSE '0' END) NUMSTJ007 
,(CASE WHEN TJ030='1' THEN (TJ047) ELSE '0' END) NUMSTJ047
,(CASE WHEN TJ030='1' THEN (TJ035) ELSE '0' END) NUMSTJ035 
,(CASE WHEN TJ030='1' THEN (TJ048) ELSE '0' END) NUMSTJ048
FROM [test0923].dbo.COPTI, [test0923].dbo.COPTJ,[test0923].dbo.INVMB
WHERE TI001=TJ001 AND TI002=TJ002
AND TJ004=MB001
AND TI001=@TI001 AND TI002=@TI002
AND ISNULL(TJ018,'')<>''

) AS TEMP 
WHERE TJ018=TD001 AND TJ019=TD002 AND TJ020=TD003


                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TI001", SqlDbType.NVarChar).Value = TI001;
                    command.Parameters.Add("@TI002", SqlDbType.NVarChar).Value = TI002;

                    command.Parameters.Add("@TJ001", SqlDbType.NVarChar).Value = TI001;
                    command.Parameters.Add("@TJ002", SqlDbType.NVarChar).Value = TI002;
                    command.Parameters.Add("@TI019", SqlDbType.NVarChar).Value = TI019;
                    command.Parameters.Add("@TI041", SqlDbType.NVarChar).Value = TI041;
                    command.Parameters.Add("@TI035", SqlDbType.NVarChar).Value = TI035;
                    command.Parameters.Add("@TJ021", SqlDbType.NVarChar).Value = TJ021;
                    command.Parameters.Add("@TJ024", SqlDbType.NVarChar).Value = TJ024;


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
