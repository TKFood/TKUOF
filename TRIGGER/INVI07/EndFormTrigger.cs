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

namespace TKUOF.TRIGGER.INVI07
{
    //INVTJ INVTK INVI07成本開帳調整單 的 核準


    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {

        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TJ001 = null;
            string TJ002 = null;
            string TJ003 = null;
            string FORMID = null;
            string MODIFIER = null;
            UserUCO userUCO = new UserUCO();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TJ001 = applyTask.Task.CurrentDocument.Fields["TJ001"].FieldValue.ToString().Trim();
            TJ002 = applyTask.Task.CurrentDocument.Fields["TJ002"].FieldValue.ToString().Trim();
            TJ003 = applyTask.Task.CurrentDocument.Fields["TJ003"].FieldValue.ToString().Trim();

            FORMID = applyTask.FormNumber;
            //MODIFIER = applyTask.Task.Applicant.Account;

            //取得簽核人工號
            EBUser ebUser = userUCO.GetEBUser(Current.UserGUID);
            MODIFIER = ebUser.Account;

           
            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(TJ001) && !string.IsNullOrEmpty(TJ002) )
                {
                    UPDATE_INVI07(TJ001, TJ002, TJ003, FORMID, MODIFIER);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATE_INVI07(string TJ001, string TJ002, string TJ003, string FORMID, string MODIFIER)
        {           
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            string TJ010 = "Y";
            string TJ013 = MODIFIER;
            string TJ015 = "N";    

            string TK001 = TJ001;
            string TK002 = TJ002;
            string TK023 = "Y";

            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"   

                                        UPDATE [test0923].dbo.INVTJ
                                        SET 
                                        TJ010=@TJ010 
                                        ,TJ013=@TJ013
                                        ,TJ015=@TJ015 
                                        ,FLAG=FLAG+1
                                        ,MODIFIER=@MODIFIER 
                                        ,MODI_DATE=@MODI_DATE
                                        ,MODI_TIME=@MODI_TIME
                                        WHERE TJ001=@TJ001 AND TJ002=@TJ002

                                        UPDATE [test0923].dbo.INVTK
                                        SET
                                        TK023=@TK023
                                        ,FLAG=FLAG+1
                                        ,MODIFIER=@MODIFIER 
                                        ,MODI_DATE=@MODI_DATE
                                        ,MODI_TIME=@MODI_TIME
                                        WHERE TK001=@TK001 AND TK002=@TK002


                                        UPDATE [test0923].dbo.INVMB
                                        SET MB064=MB064+TNUMS
                                        ,MB065=MB065+TMOMEYS
                                        ,MB089=MB089+TPACKAGES
                                        ,INVMB.FLAG=INVMB.FLAG+1
                                        FROM 
                                        (
                                        SELECT TK004,SUM(TK007*MQ010) AS 'TNUMS',SUM(TK016*MQ010) AS 'TMOMEYS',SUM(TK024*MQ010) AS 'TPACKAGES'
                                        FROM [test0923].dbo.INVTJ,[test0923].dbo.INVTK,[test0923].dbo.CMSMQ
                                        WHERE TJ001=TK001 AND TJ002=TK002
                                        AND TJ001=MQ001
                                        AND TJ001=@TJ001 AND TJ002=@TJ002
                                        GROUP BY  TK004
                                        ) AS TEMP
                                        WHERE MB001=TEMP.TK004

                                        UPDATE [test0923].dbo.INVMC
                                        SET 
                                        MC007=MC007+TNUMS
                                        ,MC008=MC008+TMOMEYS
                                        ,MC012=MC012
                                        ,MC013=MC013
                                        ,INVMC.FLAG=INVMC.FLAG+1
                                        FROM 
                                        (
                                        SELECT MQ010,TK004,SUM(TK007*MQ010) AS 'TNUMS',SUM(TK016*MQ010) AS 'TMOMEYS',TK017,TK018,SUM(TK024*MQ010) AS 'TPACKAGES'
                                        FROM [test0923].dbo.INVTJ,[test0923].dbo.INVTK,[test0923].dbo.CMSMQ
                                        WHERE TJ001=TK001 AND TJ002=TK002
                                        AND TJ001=MQ001
                                        AND TJ001=@TJ001 AND TJ002=@TJ002
                                        GROUP BY MQ010,TK004,TK017,TK018
                                        ) AS TEMP
                                        WHERE MC001=TEMP.TK004 AND MC002=TEMP.TK017

                                        INSERT INTO [test0923].dbo.INVLA 
                                        (LA001 ,LA002 , LA003 ,LA004 ,LA005 ,LA006,LA007,LA008 ,LA009 ,LA010 , 
                                        LA011 ,LA012 ,LA013 ,LA014 ,LA015 ,LA016 ,LA017,LA018,LA019,LA020,LA021, 
                                        COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
                                        CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME ) 
                                        SELECT 
                                        TK004 LA001 ,'' LA002 ,''  LA003 ,TJ003 LA004 ,MQ010 LA005 ,TK001 LA006,TK002 LA007,TK003 LA008 ,TK017 LA009 ,TK022 LA010 
                                        ,MQ010*TK007 LA011 ,(TK008+TK009+TK010+TK011) LA012 ,TK016 LA013 ,'5' LA014 ,'y' LA015 ,TK018 LA016 ,TK012 LA017,TK013 LA018,TK014 LA019,TK015 LA020,MQ010*TK024 LA021
                                        ,INVTJ.COMPANY ,INVTJ.CREATOR ,INVTJ.USR_GROUP ,INVTJ.CREATE_DATE ,0 FLAG
                                        ,INVTJ.CREATE_TIME, '' MODI_TIME, 'P004' TRANS_TYPE, 'INVI07' TRANS_NAME 
                                        FROM [test0923].dbo.INVTJ,[test0923].dbo.INVTK,[test0923].dbo.CMSMQ
                                        WHERE TJ001=TK001 AND TJ002=TK002
                                        AND TJ001=MQ001
                                        AND TJ001=@TJ001 AND TJ002=@TJ002

                                        INSERT INTO [test0923].dbo.INVME
                                        (ME001,ME002,ME003,ME004,ME005,ME006,ME007,ME009,ME010,ME032,FLAG,COMPANY,CREATOR,USR_GROUP,CREATE_DATE,
                                        CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME ) 
                                        SELECT 
                                        TK004 ME001,TK018 ME002,TJ003 ME003,'' ME004,TK001 ME005,TK002 ME006,'N' ME007,TK019 ME009,TK020 ME010,TK039 ME032
                                        ,0 FLAG,INVTJ.COMPANY,INVTJ.CREATOR,INVTJ.USR_GROUP,INVTJ.CREATE_DATE
                                        ,INVTJ.CREATE_TIME, INVTJ.MODI_TIME, 'P004' TRANS_TYPE, 'INVI07' TRANS_NAME 
                                        FROM [test0923].dbo.INVTJ,[test0923].dbo.INVTK,[test0923].dbo.CMSMQ
                                        WHERE TJ001=TK001 AND TJ002=TK002
                                        AND TJ001=MQ001
                                        AND TJ001=@TJ001 AND TJ002=@TJ002
                                        AND REPLACE(TK004+TK018,' ','') NOT IN (SELECT REPLACE(ME001+ME002,' ','') FROM [test0923].dbo.INVME )

                                        UPDATE [test0923].dbo.INVME 
                                        SET ME003=TEMP.ME003
                                        ,ME005=TEMP.ME005
                                        ,ME006=TEMP.ME006
                                        ,ME007=TEMP.ME007
                                        ,ME009=TEMP.ME009
                                        ,ME010=TEMP.ME010
                                        ,ME032=TEMP.ME032
                                        ,INVME.FLAG=INVME.FLAG+1
                                        FROM (
                                        SELECT 
                                        TK004 ME001,TK018 ME002,TJ003 ME003,'' ME004,TK001 ME005,TK002 ME006,'N' ME007,TK019 ME009,TK020 ME010,TK039 ME032
                                        ,0 FLAG,INVTJ.COMPANY,INVTJ.CREATOR,INVTJ.USR_GROUP,INVTJ.CREATE_DATE
                                        ,INVTJ.CREATE_TIME, INVTJ.MODI_TIME, 'P004' TRANS_TYPE, 'INVI07' TRANS_NAME 
                                        FROM [test0923].dbo.INVTJ,[test0923].dbo.INVTK,[test0923].dbo.CMSMQ
                                        WHERE TJ001=TK001 AND TJ002=TK002
                                        AND TJ001=MQ001
                                        AND TJ001=@TJ001 AND TJ002=@TJ002
                                        AND REPLACE(TK004+TK018,' ','')  IN (SELECT REPLACE(ME001+ME002,' ','') FROM [test0923].dbo.INVME )
                                        ) AS TEMP
                                        WHERE INVME.ME001=TEMP.ME001 AND INVME.ME002=TEMP.ME002

                                        INSERT INTO [test0923].dbo.INVMF 
                                        (MF001 ,MF002 ,MF003 ,MF004 ,MF005 ,MF006,MF007,MF008 ,MF009 ,MF010 , 
                                        MF011 ,MF012 ,MF013,MF014 ,COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
                                        CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME ) 
                                        SELECT 
                                        TK004 MF001 ,TK018 MF002 ,TJ003 MF003 ,TK001 MF004 ,TK002 MF005 ,TK003 MF006,TK017 MF007,MQ010 MF008 ,'5' MF009 ,TK007 MF010 
                                        ,'' MF011 ,'' MF012 ,'' MF013,TK024 MF014 
                                        ,INVTJ.COMPANY ,INVTJ.CREATOR ,INVTJ.USR_GROUP ,INVTJ.CREATE_DATE ,0 FLAG, 
                                        INVTJ.CREATE_TIME, INVTJ.MODI_TIME,  'P004' TRANS_TYPE, 'INVI07' TRANS_NAME 
                                        FROM [test0923].dbo.INVTJ,[test0923].dbo.INVTK,[test0923].dbo.CMSMQ
                                        WHERE TJ001=TK001 AND TJ002=TK002
                                        AND TJ001=MQ001
                                        AND TJ001=@TJ001 AND TJ002=@TJ002


                                      


                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TJ001", SqlDbType.NVarChar).Value = TJ001;
                    command.Parameters.Add("@TJ002", SqlDbType.NVarChar).Value = TJ002;

                    command.Parameters.Add("@TK001", SqlDbType.NVarChar).Value = TK001;
                    command.Parameters.Add("@TK002", SqlDbType.NVarChar).Value = TK002;
                    command.Parameters.Add("@TJ010", SqlDbType.NVarChar).Value = TJ010;
                    command.Parameters.Add("@TJ013", SqlDbType.NVarChar).Value = TJ013;
                    command.Parameters.Add("@TJ015", SqlDbType.NVarChar).Value = TJ015;
                    command.Parameters.Add("@TK023", SqlDbType.NVarChar).Value = TK023;
              

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
