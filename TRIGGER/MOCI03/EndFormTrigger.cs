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

namespace TKUOF.TRIGGER.MOCI03
{
    //ERP-MOCI03領料單 的 核準


    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {

        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TC001 = null;
            string TC002 = null;

            string FORMID = null;
            string MODIFIER = null;
            UserUCO userUCO = new UserUCO();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TC001 = applyTask.Task.CurrentDocument.Fields["TC001"].FieldValue.ToString().Trim();
            TC002 = applyTask.Task.CurrentDocument.Fields["TC002"].FieldValue.ToString().Trim();
         

            FORMID = applyTask.FormNumber;
            //MODIFIER = applyTask.Task.Applicant.Account;

            //取得簽核人工號
            EBUser ebUser = userUCO.GetEBUser(Current.UserGUID);
            MODIFIER = ebUser.Account;

           
            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002) )
                {
                    UPDATE_MOCI03(TC001, TC002, FORMID, MODIFIER);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATE_MOCI03(string TC001, string TC002, string FORMID, string MODIFIER)
        {           
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            string TC009 = "Y";
            string TC015 = MODIFIER;
            string TC016 = "N";
            string TE019 = "Y";

            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"   

                                    UPDATE [test0923].dbo.MOCTC
                                    SET 
                                    TC009=@TC009 
                                    ,TC015=@TC015 
                                    ,TC016=@TC016 
                                    ,FLAG=FLAG+1
                                    ,MODIFIER=@MODIFIER
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME 
                                    WHERE TC001=@TC001 AND TC002=@TC002

                                    UPDATE [test0923].dbo.MOCTE
                                    SET 
                                    TE019=@TE019 
                                    ,FLAG=FLAG+1
                                    ,MODIFIER=@MODIFIER
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME 
                                    WHERE TE001=@TE001 AND TE002=@TE002


                                    UPDATE [test0923].dbo.INVMB 
                                    SET MB064=MB064-TNUMS
                                    ,MB065=MB065-TMONEYS
                                    ,MB089=MB089-TPACKAGES
                                    ,FLAG=FLAG+1
                                    FROM (
                                    SELECT 
                                    TE004
                                    ,(CASE WHEN ISNULL(MD004,0)>0 THEN  SUM(TE005)*MD004 ELSE SUM(TE005) END) TNUMS
                                    ,((CASE WHEN ISNULL(MD004,0)>0 THEN  SUM(TE005)*MD004 ELSE SUM(TE005) END)*MB065/MB064) TMONEYS
                                    ,SUM(TE021) AS TPACKAGES
                                    FROM [test0923].dbo.MOCTC,[test0923].dbo.MOCTE
                                    LEFT JOIN [test0923].dbo.INVMD ON MD001=TE004 AND MD002=TE006
                                    ,[test0923].dbo.INVMB
                                    WHERE TC001=TE001 AND TC002=TE002
                                    AND MB001=TE004
                                    AND TC001=@TC001 AND TC002=@TC002
                                    GROUP BY TE004,MD004,MB065,MB064
                                    ) AS TEMP
                                    WHERE TEMP.TE004=INVMB.MB002



                                    UPDATE [test0923].dbo.INVMC 
                                    SET 
                                    MC007=MC007-TNUMS
                                    ,MC008=MC008-TMONEYS
                                    ,MC013=TC003
                                    ,MC014=MC014-TPACKAGES
                                    ,FLAG=FLAG+1
                                    FROM (
                                    SELECT 
                                    TE004,TE008,TC003
                                    ,(CASE WHEN ISNULL(MD004,0)>0 THEN  SUM(TE005)*MD004 ELSE SUM(TE005) END) TNUMS
                                    ,((CASE WHEN ISNULL(MD004,0)>0 THEN  SUM(TE005)*MD004 ELSE SUM(TE005) END)*MB065/MB064) TMONEYS
                                    ,SUM(TE021) AS TPACKAGES
                                    FROM [test0923].dbo.MOCTC,[test0923].dbo.MOCTE
                                    LEFT JOIN [test0923].dbo.INVMD ON MD001=TE004 AND MD002=TE006
                                    ,[test0923].dbo.INVMB
                                    WHERE TC001=TE001 AND TC002=TE002
                                    AND MB001=TE004
                                    AND TC001=@TC001 AND TC002=@TC002
                                    GROUP BY TE004,TE008,TC003,MD004,MB065,MB064
                                    ) AS TEMP
                                    WHERE TEMP.TE004=INVMC.MC001 AND  TEMP.TE008=INVMC.MC002



                                    INSERT INTO [test0923].dbo.INVLA 
                                    (LA001 ,LA002 , LA003 ,LA004 ,LA005 ,LA006,LA007,LA008 ,LA009 ,LA010 ,
                                     LA011 ,LA012 ,LA013 ,LA014 ,LA015 ,LA016 ,LA017,LA018,LA019,LA020,LA021, 
                                     COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
                                     CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME )
                                    SELECT
                                    TE004 LA001 ,'' LA002 ,'' LA003 ,TC003 LA004 ,-1 LA005 ,TE001 LA006,TE002 LA007,TE003 LA008 ,TE008 LA009 ,TE014 LA010 
                                    ,(CASE WHEN ISNULL(MD004,0)>0 THEN  TE005*MD004 ELSE TE005 END) LA011 ,(CASE WHEN MB065>0 AND MB064>0 THEN MB065/MB064 ELSE 0  END) LA012 ,((CASE WHEN ISNULL(MD004,0)>0 THEN  TE005*MD004 ELSE TE005 END)*((CASE WHEN MB065>0 AND MB064>0 THEN MB065/MB064 ELSE 0  END))) LA013 ,'3' LA014 ,'N' LA015 ,TE010 LA016 ,((CASE WHEN ISNULL(MD004,0)>0 THEN  TE005*MD004 ELSE TE005 END)*((CASE WHEN MB065>0 AND MB064>0 THEN MB065/MB064 ELSE 0  END))) LA017,0 LA018,0 LA019,0 LA020,0 LA021
                                    ,MOCTC.COMPANY ,MOCTC.CREATOR ,MOCTC.USR_GROUP ,MOCTC.CREATE_DATE ,0 FLAG
                                    ,MOCTC.CREATE_TIME, MOCTC.MODI_TIME,'P004' TRANS_TYPE,'MOCI03' TRANS_NAME
                                    FROM [test0923].dbo.MOCTC,[test0923].dbo.MOCTE
                                    LEFT JOIN [test0923].dbo.INVMD ON MD001=TE004 AND MD002=TE006
                                    ,[test0923].dbo.INVMB
                                    WHERE TC001=TE001 AND TC002=TE002
                                    AND MB001=TE004
                                    AND TC001=@TC001 AND TC002=@TC002

                                    INSERT INTO [test0923].dbo.INVMF 
                                    (MF001 ,MF002 ,MF003 ,MF004 ,MF005 ,MF006,MF007,MF008 ,MF009 ,MF010 ,
                                     MF011 ,MF012 ,MF013,MF014 ,COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
                                     CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME )
                                    SELECT
                                    TE004 MF001 ,TE010 MF002 ,TC003 MF003 ,TE001 MF004 ,TE002 MF005 ,TE003 MF006,TE008 MF007,-1 MF008 ,3 MF009 ,(CASE WHEN ISNULL(MD004,0)>0 THEN  TE005*MD004 ELSE TE005 END) MF010 
                                    ,'' MF011 ,'' MF012 ,'' MF013,TE021 MF014 
                                    ,MOCTC.COMPANY ,MOCTC.CREATOR ,MOCTC.USR_GROUP ,MOCTC.CREATE_DATE ,0 FLAG
                                    ,MOCTC.CREATE_TIME, MOCTC.MODI_TIME,'P004' TRANS_TYPE,'MOCI03' TRANS_NAME
                                    FROM [test0923].dbo.MOCTC,[test0923].dbo.MOCTE
                                    LEFT JOIN [test0923].dbo.INVMD ON MD001=TE004 AND MD002=TE006
                                    ,[test0923].dbo.INVMB
                                    WHERE TC001=TE001 AND TC002=TE002
                                    AND MB001=TE004
                                    AND TC001=@TC001 AND TC002=@TC002


                                    UPDATE [test0923].dbo.MOCTA
                                    SET TA011='2'
                                    ,MOCTA.FLAG=MOCTA.FLAG+1
                                    FROM 
                                    (
                                    SELECT
                                    TE011,TE012
                                    FROM [test0923].dbo.MOCTC,[test0923].dbo.MOCTE
                                    LEFT JOIN [test0923].dbo.INVMD ON MD001=TE004 AND MD002=TE006
                                    ,[test0923].dbo.INVMB
                                    WHERE TC001=TE001 AND TC002=TE002
                                    AND MB001=TE004
                                    AND TC001=@TC001 AND TC002=@TC002
                                    GROUP BY TE011,TE012
                                    ) AS TEMP
                                    WHERE TEMP.TE011=TA001 AND TEMP.TE012=TA002


                                    UPDATE [test0923].dbo.MOCTB 
                                    SET TB005=TE005
                                    ,TB016=TC003
                                    ,TB020=TE021
                                    FROM (
                                    SELECT
                                    TE011,TE012,TE004,TE005,TE009,TC003,TE021
                                    FROM [test0923].dbo.MOCTC,[test0923].dbo.MOCTE
                                    LEFT JOIN [test0923].dbo.INVMD ON MD001=TE004 AND MD002=TE006
                                    ,[test0923].dbo.INVMB
                                    WHERE TC001=TE001 AND TC002=TE002
                                    AND MB001=TE004
                                    AND TC001=@TC001 AND TC002=@TC002
                                    ) AS TEMP
                                    WHERE TEMP.TE011=TB001 AND TEMP.TE012=TB002  AND TEMP.TE004=TB003 AND TEMP.TE009=TB006
                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TC001", SqlDbType.NVarChar).Value = TC001;
                    command.Parameters.Add("@TC002", SqlDbType.NVarChar).Value = TC002;
                    command.Parameters.Add("@TE001", SqlDbType.NVarChar).Value = TC001;
                    command.Parameters.Add("@TE002", SqlDbType.NVarChar).Value = TC002;

                    command.Parameters.Add("@TC009", SqlDbType.NVarChar).Value = TC009;
                    command.Parameters.Add("@TC015", SqlDbType.NVarChar).Value = TC015;
                    command.Parameters.Add("@TC016", SqlDbType.NVarChar).Value = TC016;
                    command.Parameters.Add("@TE019", SqlDbType.NVarChar).Value = TE019;

                
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
