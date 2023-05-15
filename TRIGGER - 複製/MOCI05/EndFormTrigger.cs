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

namespace TKUOF.TRIGGER.MOCI05
{
    //ERP-MOCI05.生產入庫單 的 核準


    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {

        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TF001 = null;
            string TF002 = null;
            string TF003 = null;
            string FORMID = null;
            string MODIFIER = null;
            UserUCO userUCO = new UserUCO();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TF001 = applyTask.Task.CurrentDocument.Fields["TF001"].FieldValue.ToString().Trim();
            TF002 = applyTask.Task.CurrentDocument.Fields["TF002"].FieldValue.ToString().Trim();
            TF003 = applyTask.Task.CurrentDocument.Fields["TF003"].FieldValue.ToString().Trim();

            FORMID = applyTask.FormNumber;
            //MODIFIER = applyTask.Task.Applicant.Account;

            //取得簽核人工號
            EBUser ebUser = userUCO.GetEBUser(Current.UserGUID);
            MODIFIER = ebUser.Account;

           
            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(TF001) && !string.IsNullOrEmpty(TF002) )
                {
                    UPDATE_MOCI05(TF001, TF002, TF003, FORMID, MODIFIER);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATE_MOCI05(string TF001, string TF002, string TF003, string FORMID, string MODIFIER)
        {           
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            string TF006 = "Y";
            string TF013 = MODIFIER;
            string TF014 = "N";
            string TG022 = "Y";

            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"   

                                        UPDATE  [test0923].dbo.MOCTF
                                        SET
                                        TF006=@TF006 
                                        ,TF013=@TF013  
                                        ,TF014=@TF014 
                                        ,FLAG=FLAG+1
                                        ,MODIFIER=@MODIFIER 
                                        ,MODI_DATE=@MODI_DATE
                                        ,MODI_TIME=@MODI_TIME 
                                        WHERE TF001=@TF001 AND TF002=@TF002

                                        UPDATE [test0923].dbo.MOCTG
                                        SET
                                        TG022=@TG022 
                                        ,FLAG=FLAG+1
                                        ,MODIFIER=@MODIFIER
                                        ,MODI_DATE=@MODI_DATE
                                        ,MODI_TIME=@MODI_TIME 
                                        WHERE TG001=@TG001 AND TG002=@TG002

                                        UPDATE [test0923].dbo.INVMB
                                        SET MB064=MB064+TUMS
                                        ,MB065=MB065+TMONEYS
                                        ,MB089=MB089+TPACKAGES
                                        ,INVMB.FLAG=INVMB.FLAG+1
                                        FROM 
                                        (
                                        SELECT TG004,SUM(TG011) TUMS,SUM(TG025) TPACKAGES,((CASE WHEN MB065>0 AND MB064>0 THEN (MB065/MB064) ELSE 0 END )*SUM(TG011)) TMONEYS 
                                        FROM  [test0923].dbo.MOCTF,[test0923].dbo.MOCTG,[test0923].dbo.INVMB
                                        WHERE TF001=TG001 AND TF002=TG002
                                        AND MB001=TG004
                                        AND TF001=@TF001 AND TF002=@TF002
                                        GROUP BY TG004,MB065,MB064
                                        ) AS TEMP
                                        WHERE TEMP.TG004=MB001



                                        UPDATE [test0923].dbo.INVMC
                                        SET  MC007=MC007+TUMS
                                        ,MC008=MC008+TMONEYS
                                        ,MC012=@TF003
                                        ,MC014=MC014+TPACKAGES
                                        ,INVMC.FLAG=INVMC.FLAG+1 
                                        FROM 
                                        (
                                        SELECT TG004,TG010,SUM(TG011) TUMS,SUM(TG025) TPACKAGES,((CASE WHEN MB065>0 AND MB064>0 THEN (MB065/MB064) ELSE 0 END )*SUM(TG011)) TMONEYS 
                                        FROM  [test0923].dbo.MOCTF,[test0923].dbo.MOCTG,[test0923].dbo.INVMB
                                        WHERE TF001=TG001 AND TF002=TG002
                                        AND MB001=TG004
                                        AND TF001=@TF001 AND TF002=@TF002
                                        GROUP BY TG004,TG010,MB065,MB064
                                        ) AS TEMP 
                                        WHERE TEMP.TG004=MC001 AND TEMP.TG010=MC002



                                        INSERT INTO [test0923].dbo.INVLA (
                                        LA001 ,LA002 , LA003 ,LA004 ,LA005 ,LA006,LA007,LA008 ,LA009 ,LA010 , 
                                        LA011 ,LA012 ,LA013 ,LA014 ,LA015 ,LA016 ,LA017,LA018,LA019,LA020,LA021, 
                                        COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
                                        CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME ) 
                                        SELECT TG004 LA001 ,'' LA002 ,'' LA003 ,TF003 LA004 ,'1' LA005 ,TG001 LA006,TG002 LA007,TG003 LA008 ,TG010 LA009 ,TG011 LA010 , 
                                        (CASE WHEN ISNULL(MD004,0)>0 THEN  TG011*MD004 ELSE TG011 END) LA011 ,((CASE WHEN MB065>0 AND MB064>0 THEN (MB065/MB064) ELSE 0 END ))  LA012 ,((CASE WHEN MB065>0 AND MB064>0 THEN (MB065/MB064) ELSE 0 END )*((CASE WHEN ISNULL(MD004,0)>0 THEN  TG011*MD004 ELSE TG011 END)))  LA013 ,'1' LA014 ,'Y' LA015 ,TG017 LA016 ,((CASE WHEN MB065>0 AND MB064>0 THEN (MB065/MB064) ELSE 0 END )*(TG011))  LA017,0 LA018, 0 LA019,0 LA020, 0 LA021, 
                                        MOCTF.COMPANY ,MOCTF.CREATOR ,MOCTF.USR_GROUP ,MOCTF.CREATE_DATE ,0 FLAG, 
                                        MOCTF.CREATE_TIME, MOCTF.MODI_TIME, 'P004' TRANS_TYPE, 'MOCI05' TRANS_NAME
                                        FROM [test0923].dbo.MOCTF
                                        JOIN [test0923].dbo.MOCTG ON TF001=TG001 AND TF002=TG002
                                        JOIN [test0923].dbo.INVMB ON MB001=TG004
                                        LEFT JOIN [test0923].dbo.INVMD ON MD001=MB001 AND MD002=TG007
                                        JOIN [test0923].dbo.CMSMQ ON TG001=MQ001
                                        WHERE  TF001=@TF001 AND TF002=@TF002

                                        INSERT INTO   [test0923].dbo.INVME (
                                        ME001,ME002,ME003,ME004,ME005,ME006,ME007,ME008,ME009,ME010,ME032,FLAG,COMPANY,CREATOR,USR_GROUP,CREATE_DATE,
                                        CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME ) 
                                        SELECT TG004 ME001,TG017 ME002,TF003 ME003,'' ME004,TG001 ME005,TG002 ME006,'N' ME007,'' ME008,TG018 ME009,TG019 ME010,TF003 ME032
                                        ,0 FLAG,MOCTF.COMPANY,MOCTF.CREATOR,MOCTF.USR_GROUP,MOCTF.CREATE_DATE
                                        ,MOCTF.CREATE_TIME, MOCTF.MODI_TIME, 'P004' TRANS_TYPE, 'MOCI05' TRANS_NAME
                                        FROM [test0923].dbo.MOCTF
                                        JOIN [test0923].dbo.MOCTG ON TF001=TG001 AND TF002=TG002
                                        JOIN [test0923].dbo.INVMB ON MB001=TG004
                                        LEFT JOIN [test0923].dbo.INVMD ON MD001=MB001 AND MD002=TG007
                                        JOIN [test0923].dbo.CMSMQ ON TG001=MQ001
                                        WHERE TF001=@TF001 AND TF002=@TF002
                                        AND REPLACE(TG004+TG017,' ','') NOT IN (SELECT REPLACE(ME001+ME002,' ','') FROM [test0923].dbo.INVME )

                                        UPDATE [test0923].dbo.INVME
                                        SET  ME003=TEMP.ME003
                                        ,ME004=''
                                        ,ME005=TEMP.ME005 
                                        ,ME006=TEMP.ME006
                                        ,ME007='N'
                                        ,ME008=''
                                        ,ME009=TEMP.ME009
                                        ,ME010=TEMP.ME010
                                        ,ME032=TEMP.ME032
                                        ,INVME.FLAG=INVME.FLAG+1
                                        FROM (
                                        SELECT TG004 ME001,TG017 ME002,TF003 ME003,'' ME004,TG001 ME005,TG002 ME006,'N' ME007,'' ME008,TG018 ME009,TG019 ME010,TF003 ME032
                                        ,0 FLAG,MOCTF.COMPANY,MOCTF.CREATOR,MOCTF.USR_GROUP,MOCTF.CREATE_DATE
                                        ,MOCTF.CREATE_TIME, MOCTF.MODI_TIME, 'P004' TRANS_TYPE, 'MOCI05' TRANS_NAME
                                        FROM [test0923].dbo.MOCTF
                                        JOIN [test0923].dbo.MOCTG ON TF001=TG001 AND TF002=TG002
                                        JOIN [test0923].dbo.INVMB ON MB001=TG004
                                        LEFT JOIN [test0923].dbo.INVMD ON MD001=MB001 AND MD002=TG007
                                        JOIN [test0923].dbo.CMSMQ ON TG001=MQ001
                                        WHERE TF001=@TF001 AND TF002=@TF002
                                        AND REPLACE(TG004+TG017,' ','') IN (SELECT REPLACE(ME001+ME002,' ','') FROM [test0923].dbo.INVME )
                                        ) AS TEMP 
                                        WHERE TEMP.ME001=INVME.ME001 AND TEMP.ME002=INVME.ME002 

                                        INSERT INTO [test0923].dbo.INVMF 
                                        (MF001 ,MF002 ,MF003 ,MF004 ,MF005 ,MF006,MF007,MF008 ,MF009 ,MF010 , 
                                        MF011 ,MF012 ,MF013,MF014 ,COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
                                        CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME )
                                        SELECT TG004 MF001 ,TG017 MF002 ,TF003 MF003 ,TF001 MF004 ,TF002 MF005 ,TG003 MF006,TG010 MF007,'1' MF008 ,'1' MF009 ,TG011 MF010 
                                        ,'' MF011 ,'' MF012 ,'' MF013,TG025 MF014 
                                        ,MOCTF.COMPANY ,MOCTF.CREATOR ,MOCTF.USR_GROUP ,MOCTF.CREATE_DATE ,0 FLAG
                                        ,MOCTF.CREATE_TIME, MOCTF.MODI_TIME,'P004' TRANS_TYPE,'MOCI05' TRANS_NAME 
                                        FROM [test0923].dbo.MOCTF
                                        JOIN [test0923].dbo.MOCTG ON TF001=TG001 AND TF002=TG002
                                        JOIN [test0923].dbo.INVMB ON MB001=TG004
                                        LEFT JOIN [test0923].dbo.INVMD ON MD001=MB001 AND MD002=TG007
                                        JOIN [test0923].dbo.CMSMQ ON TG001=MQ001
                                        WHERE TF001=@TF001 AND TF002=@TF002

                                        UPDATE  [test0923].dbo.MOCTA 
                                        SET TA011=(CASE WHEN TA015=TG011 THEN 'Y' ELSE '3' END)
                                        ,TA012=TF003
                                        ,TA014=(CASE WHEN TA015=TG011 THEN TF003 ELSE '' END) 
                                        ,TA017=TA017+TG011
                                        ,TA018=TG012
                                        ,TA046=TG025
                                        ,TA047=TG026
                                        ,MOCTA.FLAG=MOCTA.FLAG+1
                                        ,MODIFIER=@MODIFIER
                                        ,MODI_DATE=@MODI_DATE
                                        ,MODI_TIME=@MODI_TIME
                                        FROM (
                                        SELECT TG001,TG002,TG003,TG014,TG015,TF003,TG011,TG012,TG025,TG026
                                        FROM [test0923].dbo.MOCTF
                                        JOIN [test0923].dbo.MOCTG ON TF001=TG001 AND TF002=TG002
                                        JOIN [test0923].dbo.INVMB ON MB001=TG004
                                        LEFT JOIN [test0923].dbo.INVMD ON MD001=MB001 AND MD002=TG007
                                        JOIN [test0923].dbo.CMSMQ ON TG001=MQ001
                                        WHERE TF001=@TF001 AND TF002=@TF002
                                        ) AS TEMP 
                                        WHERE TEMP.TG014=TA001 AND TEMP.TG015=TA001
                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TF001", SqlDbType.NVarChar).Value = TF001;
                    command.Parameters.Add("@TF002", SqlDbType.NVarChar).Value = TF002;
                    command.Parameters.Add("@TF003", SqlDbType.NVarChar).Value = TF003;
                    command.Parameters.Add("@TG001", SqlDbType.NVarChar).Value = TF001;
                    command.Parameters.Add("@TG002", SqlDbType.NVarChar).Value = TF002;

                    command.Parameters.Add("@TF006", SqlDbType.NVarChar).Value = TF006;
                    command.Parameters.Add("@TF013", SqlDbType.NVarChar).Value = TF013;
                    command.Parameters.Add("@TF014", SqlDbType.NVarChar).Value = TF014;
                    command.Parameters.Add("@TG022", SqlDbType.NVarChar).Value = TG022;
             

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
