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

namespace TKUOF.TRIGGER.COP30
{
    //核準


    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {

        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TG001 = null;
            string TG002 = null;
            string FORMID = null;
            string MODIFIER = null;
            UserUCO userUCO = new UserUCO();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TG001 = applyTask.Task.CurrentDocument.Fields["TG001"].FieldValue.ToString().Trim();
            TG002 = applyTask.Task.CurrentDocument.Fields["TG002"].FieldValue.ToString().Trim();

            FORMID = applyTask.FormNumber;
            //MODIFIER = applyTask.Task.Applicant.Account;

            //取得簽核人工號
            EBUser ebUser = userUCO.GetEBUser(Current.UserGUID);
            MODIFIER = ebUser.Account;

           
            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(TG001) && !string.IsNullOrEmpty(TG002) )
                {
                    UPDATE_COP30(TG001, TG002, FORMID, MODIFIER);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATE_COP30(string TG001, string TG002, string FORMID, string MODIFIER)
        {           
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            string TG023 = "Y";
            string TG043 = MODIFIER;
            string TH020 = "Y";
            string TH026 = "N";

            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"       
 UPDATE [test0923].dbo.COPTG
SET
TG023=@TG023 
,TG043=@TG043  
,FLAG=FLAG+1
,MODIFIER=@MODIFIER
,MODI_DATE=@MODI_DATE
,MODI_TIME=@MODI_TIME 
WHERE TG001=@TG001 AND TG002=@TG002

UPDATE [test0923].dbo.COPTH
SET
TH020=@TH020 
,TH026=@TH026 
,FLAG=FLAG+1
,MODIFIER=@MODIFIER
,MODI_DATE=@MODI_DATE
,MODI_TIME=@MODI_TIME 
WHERE TH001=@TH001 AND TH002=@TH002

UPDATE [test0923].dbo.INVMB
SET  MB064=MB064-NUMS, MB065=MB065-MONEYS,MB089=MB089-PNUMS
FROM 
(
SELECT TG003,TH004,SUM(TH008+TH024) AS NUMS,(CASE WHEN MB065>0 AND MB064>0 THEN SUM(TH008+TH024)*(MB065/MB064) ELSE 0 END) AS 'MONEYS',SUM(TH039+TH040) AS 'PNUMS'
FROM [test0923].dbo.COPTG,[test0923].dbo.COPTH,[test0923].dbo.INVMB
WHERE TG001=TH001 AND TG002=TH002
AND TH004=MB001
AND TG001=@TG001 AND TG002=@TG002
GROUP BY TG003,TH004,MB064,MB065
) AS TEMP
WHERE TEMP.TH004=INVMB.MB001

UPDATE [test0923].dbo.INVMC
SET   MC007=MC007-NUMS,MC008=MC008-MONEYS,MC013=TG003,MC014=MC014-PNUMS
FROM 
(
SELECT TG003,TH004,SUM(TH008+TH024) AS NUMS,(CASE WHEN MB065>0 AND MB064>0 THEN SUM(TH008+TH024)*(MB065/MB064) ELSE 0 END) AS 'MONEYS',SUM(TH039+TH040) AS 'PNUMS',TH007
FROM [test0923].dbo.COPTG,[test0923].dbo.COPTH,[test0923].dbo.INVMB
WHERE TG001=TH001 AND TG002=TH002
AND TH004=MB001
AND TG001=@TG001 AND TG002=@TG002
GROUP BY TG003,TH004,MB064,MB065,TH007
) AS TEMP
WHERE TEMP.TH004=INVMC.MC001 AND TEMP.TH007=INVMC.MC002

UPDATE [test0923].dbo.COPMA
SET  MA022=TG003
FROM 
(
SELECT TG003,TG004
FROM [test0923].dbo.COPTG,[test0923].dbo.COPTH,[test0923].dbo.INVMB
WHERE TG001=TH001 AND TG002=TH002
AND TH004=MB001
AND TG001=@TG001 AND TG002=@TG002
GROUP BY TG003,TG004
) AS TEMP
WHERE TEMP.TG004=COPMA.MA001

UPDATE [test0923].dbo.COPTD
SET  TD009=TD009+TH008,TD025=TH024,TD033=TH039,TD035=TH040 ,FLAG=FLAG+1 
,TD016= (CASE WHEN TD008=TD009+TH008 THEN 'Y' ELSE 'N' END )
FROM 
(
SELECT TG003,TH004,SUM(TH008) TH008,SUM(TH024) AS TH024,SUM(TH039) TH039,SUM(TH040) TH040 ,TH014,TH015,TH016
FROM [test0923].dbo.COPTG,[test0923].dbo.COPTH,[test0923].dbo.INVMB
WHERE TG001=TH001 AND TG002=TH002
AND TH004=MB001
AND TG001=@TG001 AND TG002=@TG002
GROUP BY TG003,TH004,TH014,TH015,TH016
) AS TEMP
WHERE TEMP.TH014=TD001 AND TEMP.TH015=TD003 AND TEMP.TH016=TD003 

INSERT INTO [test0923].dbo.INVLA 
(LA001 ,LA002 , LA003 ,LA004 ,LA005 ,LA006,LA007,LA008 ,LA009 ,LA010 , 
LA011 ,LA012 ,LA013 ,LA014 ,LA015 ,LA016 ,LA017,LA018,LA019,LA020,LA021, 
COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME )
SELECT  TH004 LA001 ,'' LA002 ,'' LA003 ,TG003 LA004 ,-1 LA005 ,TH001 LA006,TH002 LA007,TH003 LA008 ,TH007 LA009 ,TH018 LA010 , 
(TH008+TH024) LA011 ,(CASE WHEN MB065>0 AND MB064>0 THEN (MB065/MB064) ELSE 0 END) LA012 ,(CASE WHEN MB065>0 AND MB064>0 THEN (TH008+TH024)*(MB065/MB064) ELSE 0 END) LA013 ,'2' LA014 ,'N' LA015 ,TH017 LA016 ,(CASE WHEN MB065>0 AND MB064>0 THEN (TH008+TH024)*(MB065/MB064) ELSE 0 END)  LA017,0 LA018,0 LA019,0 LA020,0 LA021, 
COPTG.COMPANY ,COPTG.CREATOR ,COPTG.USR_GROUP ,COPTG.CREATE_DATE ,0 FLAG, 
COPTG.CREATE_TIME, COPTG.MODI_TIME, COPTG.TRANS_TYPE, COPTG.TRANS_NAME
FROM [test0923].dbo.COPTG,[test0923].dbo.COPTH,[test0923].dbo.INVMB
WHERE TG001=TH001 AND TG002=TH002
AND TH004=MB001
AND TG001=@TG001 AND TG002=@TG002

INSERT INTO [test0923].dbo.INVMF 
(MF001 ,MF002 ,MF003 ,MF004 ,MF005 ,MF006,MF007,MF008 ,MF009 ,MF010 , 
MF011 ,MF012 ,MF013,MF014 ,COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME ) 
SELECT TH004  MF001 ,TH017 MF002 ,TG003 MF003 ,TH001 MF004 ,TH002 MF005 ,TH003 MF006,TH007 MF007,-1 MF008 ,'2' MF009 ,(TH008+TH024) MF010 , 
'' MF011 ,'' MF012 ,TH018 MF013,(TH039+TH040) MF014 
,COPTG.COMPANY ,COPTG.CREATOR ,COPTG.USR_GROUP ,COPTG.CREATE_DATE ,0 FLAG, 
COPTG.CREATE_TIME, COPTG.MODI_TIME, COPTG.TRANS_TYPE, COPTG.TRANS_NAME 
FROM [test0923].dbo.COPTG,[test0923].dbo.COPTH,[test0923].dbo.INVMB
WHERE TG001=TH001 AND TG002=TH002
AND TH004=MB001
AND TG001=@TG001 AND TG002=@TG002


                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TG001", SqlDbType.NVarChar).Value = TG001;
                    command.Parameters.Add("@TG002", SqlDbType.NVarChar).Value = TG002;

                    command.Parameters.Add("@TH001", SqlDbType.NVarChar).Value = TG001;
                    command.Parameters.Add("@TH002", SqlDbType.NVarChar).Value = TG002;

                    command.Parameters.Add("@TG023", SqlDbType.NVarChar).Value = TG023;
                    command.Parameters.Add("@TG043", SqlDbType.NVarChar).Value = TG043;
                    command.Parameters.Add("@TH020", SqlDbType.NVarChar).Value = TH020;
                    command.Parameters.Add("@TH026", SqlDbType.NVarChar).Value = TH026;

    

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
