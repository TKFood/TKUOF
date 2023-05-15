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

namespace TKUOF.TRIGGER.MOCI06
{
    //MOCTH MOCTI MOCI06.託外進貨單 的 核準


    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {

        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TH001 = null;
            string TH002 = null;
            string TH003 = null;
            string FORMID = null;
            string MODIFIER = null;
            UserUCO userUCO = new UserUCO();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TH001 = applyTask.Task.CurrentDocument.Fields["TH001"].FieldValue.ToString().Trim();
            TH002 = applyTask.Task.CurrentDocument.Fields["TH002"].FieldValue.ToString().Trim();
            TH003 = applyTask.Task.CurrentDocument.Fields["TH003"].FieldValue.ToString().Trim();

            FORMID = applyTask.FormNumber;
            //MODIFIER = applyTask.Task.Applicant.Account;

            //取得簽核人工號
            EBUser ebUser = userUCO.GetEBUser(Current.UserGUID);
            MODIFIER = ebUser.Account;

           
            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(TH001) && !string.IsNullOrEmpty(TH002) )
                {
                    UPDATE_MOCI06(TH001, TH002, TH003, FORMID, MODIFIER);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATE_MOCI06(string TH001, string TH002, string TH003, string FORMID, string MODIFIER)
        {           
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            string TH023 = "Y";
            string TH036 = "N";
            string TH046 = "MOCI06";

            string TI001 = TH001;
            string TI002 = TH002;
            string TI037 = "Y";
            string TI043 = MODIFIER;
            string TI053 = "N";

            string MC012 = TH003;

            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"   


                                    UPDATE [test0923].dbo.MOCTH
                                    SET 
                                    TH023=@TH023 
                                    ,TH036=@TH036
                                    ,TH046=@TH046 
                                    ,FLAG=FLAG+1
                                    ,MODIFIER=@MODIFIER
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME
                                    WHERE TH001=@TH001 AND TH002=@TH002
                                      

                                    UPDATE [test0923].dbo.MOCTI
                                    SET 
                                    TI037=@TI037 
                                    ,TI038='Y' 
                                    ,TI043=@TI043 
                                    ,TI053=@TI053
                                    ,FLAG=FLAG+1
                                    ,MODIFIER=@MODIFIER 
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME 
                                    WHERE TI001=@TI001 AND TI002=@TI002


                                    UPDATE [test0923].dbo.INVMB
                                     SET 
                                     MB064=MB064+TNUMS
                                    ,MB065=MB065+TMONEYS
                                    ,MB089=MB069+TPACKAGES
                                    ,INVMB.FLAG=INVMB.FLAG+1
                                    FROM (
                                    SELECT TI004,SUM(TI007) TNUMS,(CASE WHEN MB064>0 AND MB065>0 THEN SUM(TI007)*MB065/MB064 ELSE 0 END ) TMONEYS,SUM(TI016) AS TPACKAGES
                                    FROM [test0923].dbo.MOCTH, [test0923].dbo.MOCTI, [test0923].dbo.INVMB
                                    WHERE TH001=TI001 AND TH002=TI002
                                    AND TI004=MB001
                                    AND TH001=@TH001  AND TH002=@TH002
                                    GROUP BY TI004,MB064,MB065
                                    )AS TEMP
                                    WHERE TEMP.TI004=MB001

                                    UPDATE [test0923].dbo.INVMC
                                    SET 
                                    MC007=MC007+TNUMS
                                    ,MC008=MC008+TMONEYS
                                    ,MC012=@MC012
                                    ,INVMC.FLAG=INVMC.FLAG+1
                                    FROM 
                                    (
                                    SELECT TI004,TI009,SUM(TI007) TNUMS,(CASE WHEN MB064>0 AND MB065>0 THEN SUM(TI007)*MB065/MB064 ELSE 0 END ) TMONEYS,SUM(TI016) AS TPACKAGES
                                    FROM [test0923].dbo.MOCTH, [test0923].dbo.MOCTI, [test0923].dbo.INVMB
                                    WHERE TH001=TI001 AND TH002=TI002
                                    AND TI004=MB001
                                    AND TH001=@TH001  AND TH002=@TH002
                                    GROUP BY TI004,TI009,MB064,MB065
                                    )AS TEMP
                                    WHERE TEMP.TI004=MC001 AND TEMP.TI009=MC002


                                    INSERT INTO [test0923].dbo.INVLA 
                                    (LA001 ,LA002 , LA003 ,LA004 ,LA005 ,LA006,LA007,LA008 ,LA009 ,LA010 , 
                                     LA011 ,LA012 ,LA013 ,LA014 ,LA015 ,LA016 ,LA017,LA018,LA019,LA020,LA021, 
                                     COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
                                     CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME ) 
                                    SELECT TI004 LA001 ,'' LA002 ,'' LA003 ,TH003 LA004 ,MQ010 LA005 ,TI001 LA006,TI002 LA007,TI003 LA008 ,TI009 LA009 ,TI040 LA010 
                                    ,TI007 LA011 ,(CASE WHEN MB064>0 AND MB065>0 THEN MB065/MB064 ELSE 0 END ) LA012 ,TI046 LA013 ,MQ008 LA014 ,'Y' LA015 ,TI010 LA016 ,(CASE WHEN MB064>0 AND MB065>0 THEN MB065/MB064*TI007 ELSE 0 END )LA017,0 LA018,0 LA019,TI046 LA020,TI016 LA021
                                    ,MOCTH.COMPANY ,MOCTH.CREATOR ,MOCTH.USR_GROUP ,MOCTH.CREATE_DATE ,0 FLAG
                                    ,MOCTH.CREATE_TIME, MOCTH.MODI_TIME, 'P004' TRANS_TYPE,'MOCI06' TRANS_NAME
                                    FROM [test0923].dbo.MOCTH, [test0923].dbo.MOCTI,[test0923].dbo.CMSMQ,[test0923].dbo.INVMB
                                    WHERE TH001=TI001 AND TH002=TI002
                                    AND TI001=MQ001
                                    AND TI004=MB001
                                    AND TH001=@TH001  AND TH002=@TH002


                                    UPDATE [test0923].dbo.INVME
                                    SET 
                                    ME003=TH003
                                    ,ME004=''
                                    ,ME005=TH001 
                                    ,ME006=TH002
                                    ,ME007='N'
                                    ,ME008=TH001+TH002
                                    ,ME009=ME009
                                    ,ME010=ME010
                                    ,ME032=TH003
                                    ,INVME.FLAG=INVME.FLAG+1
                                    FROM (
                                    SELECT TH003,TH001,TH002,TI004,TI010
                                    FROM [test0923].dbo.MOCTH, [test0923].dbo.MOCTI,[test0923].dbo.CMSMQ,[test0923].dbo.INVMB
                                    WHERE TH001=TI001 AND TH002=TI002
                                    AND TI001=MQ001
                                    AND TI004=MB001
                                    AND TH001=@TH001  AND TH002=@TH002
                                    )AS TEMP
                                    WHERE TEMP.TI004=ME001 AND TEMP.TI010=ME002


                                    INSERT INTO [test0923].dbo.INVMF 
                                    (MF001 ,MF002 ,MF003 ,MF004 ,MF005 ,MF006,MF007,MF008 ,MF009 ,MF010 , 
                                     MF011 ,MF012 ,MF013,MF014 ,COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
                                     CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME ) 
                                    SELECT TI004 MF001 ,TI010 MF002 ,TH003 MF003 ,TI001 MF004 ,TI002 MF005 ,TI003 MF006,TI009 MF007,MQ010 MF008 ,MQ008 MF009 ,TI007 MF010 
                                    ,'' MF011 ,'' MF012 ,TH001+TH002 MF013,0 MF014 
                                    ,MOCTH.COMPANY ,MOCTH.CREATOR ,MOCTH.USR_GROUP ,MOCTH.CREATE_DATE ,0 FLAG
                                    ,MOCTH.CREATE_TIME, MOCTH.MODI_TIME, 'P004' TRANS_TYPE,'MOCI06' TRANS_NAME
                                    FROM [test0923].dbo.MOCTH, [test0923].dbo.MOCTI,[test0923].dbo.CMSMQ,[test0923].dbo.INVMB
                                    WHERE TH001=TI001 AND TH002=TI002
                                    AND TI001=MQ001
                                    AND TI004=MB001
                                    AND TH001=@TH001  AND TH002=@TH002
                                    AND REPLACE(TI004+TI010,' ','') NOT IN (SELECT REPLACE(MF001+MF002,' ','') FROM [test0923].dbo.INVMF )


                                    UPDATE [test0923].dbo.MOCTA
                                    SET TA011=(CASE WHEN TA015=TI007 THEN 'Y' ELSE 'N' END )
                                    ,TA012=TH003
                                    ,TA014=(CASE WHEN TA015=TI007 THEN TH003 ELSE '' END )
                                    ,TA017=TA017+TI007
                                    ,TA018=0
                                    ,TA046=TI016
                                    ,TA047=0 
                                    ,MOCTA.FLAG=MOCTA.FLAG+1
                                    FROM 
                                    (
                                    SELECT TH003,TH001,TH002,TI004,TI010,TI007,TI013,TI014,TI016
                                    FROM [test0923].dbo.MOCTH, [test0923].dbo.MOCTI,[test0923].dbo.CMSMQ,[test0923].dbo.INVMB
                                    WHERE TH001=TI001 AND TH002=TI002
                                    AND TI001=MQ001
                                    AND TI004=MB001
                                    AND TH001=@TH001  AND TH002=@TH002
                                    )AS TEMP
                                    WHERE TEMP.TI013=TA001 AND TEMP.TI014=TA002


                                    UPDATE [test0923].dbo.PURMA
                                    SET
                                    MA023=TH003
                                    FROM 
                                    (
                                    SELECT TH005,TH003,TH001,TH002,TI004,TI010,TI007,TI013,TI014,TI016
                                    FROM [test0923].dbo.MOCTH, [test0923].dbo.MOCTI,[test0923].dbo.CMSMQ,[test0923].dbo.INVMB
                                    WHERE TH001=TI001 AND TH002=TI002
                                    AND TI001=MQ001
                                    AND TI004=MB001
                                    AND TH001=@TH001  AND TH002=@TH002
                                    )AS TEMP
                                    WHERE PURMA.MA001=TEMP.TH005

                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TH001", SqlDbType.NVarChar).Value = TH001;
                    command.Parameters.Add("@TH002", SqlDbType.NVarChar).Value = TH002;

                    command.Parameters.Add("@TI001", SqlDbType.NVarChar).Value = TH001;
                    command.Parameters.Add("@TI002", SqlDbType.NVarChar).Value = TH002;

                    command.Parameters.Add("@TH023", SqlDbType.NVarChar).Value = TH023;
                    command.Parameters.Add("@TH036", SqlDbType.NVarChar).Value = TH036;
                    command.Parameters.Add("@TH046", SqlDbType.NVarChar).Value = TH046;
                    command.Parameters.Add("@TI037", SqlDbType.NVarChar).Value = TI037;
                    command.Parameters.Add("@TI043", SqlDbType.NVarChar).Value = TI043;
                    command.Parameters.Add("@TI053", SqlDbType.NVarChar).Value = TI053;
                    command.Parameters.Add("@MC012", SqlDbType.NVarChar).Value = MC012;

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
