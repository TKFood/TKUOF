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

namespace TKUOF.TRIGGER.COPTAB
{
    //訂單的核準

    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {

        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TA001 = null;
            string TA002 = null;
            string MB009 = null;
            string FORMID = null;
            string MODIFIER = null;
     
            UserUCO userUCO = new UserUCO();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TA001 = applyTask.Task.CurrentDocument.Fields["TA001"].FieldValue.ToString().Trim();
            TA002 = applyTask.Task.CurrentDocument.Fields["TA002"].FieldValue.ToString().Trim();
            MB009 = applyTask.Task.CurrentDocument.Fields["TA003"].FieldValue.ToString().Trim();
            FORMID = applyTask.FormNumber;
            //MODIFIER = applyTask.Task.Applicant.Account;

            //取得簽核人工號
            EBUser ebUser = userUCO.GetEBUser(Current.UserGUID);          
            MODIFIER = ebUser.Account;
            //MODIFIER = "160115";




            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(TA001) && !string.IsNullOrEmpty(TA002))
                {
                    UPDATECOPTAB(TA001, TA002, FORMID, MODIFIER,MB009);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATECOPTAB(string TA001, string TA002, string FORMID, string MODIFIER,string MB009)
        {
            string TC027 = "Y";
            string TC048 = "N";
            string TD021 = "Y";
            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
            string MB012 = "報價單"+ TA001 + '-' + TA002;

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();


         

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                        UPDATE [TK].dbo.COPTB
                                        SET TB011=@TB011
                                        ,FLAG=FLAG+1,MODIFIER=@MODIFIER,MODI_DATE=@MODI_DATE,MODI_TIME=@MODI_TIME
                                        WHERE TB001=@TB001 AND TB002=@TB002

                                        UPDATE [TK].dbo.COPTA
                                        SET TA016=@TA016,TA019=@TA019,TA029=@TA029
                                        ,FLAG=FLAG+1,MODIFIER=@MODIFIER,MODI_DATE=@MODI_DATE,MODI_TIME=@MODI_TIME
                                        ,UDF02=@UDF02
                                        WHERE TA001=@TA001 AND TA002=@TA002

                                        INSERT INTO [TK].dbo.COPMB
                                        (MB001, MB002, MB003, MB004, MB005, MB007, MB008, MB009,  MB010, MB012, MB013, MB014, MB015, MB016, MB017, MB018,  COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG) 
                                        SELECT TA004 MB001,TB004 MB002,TB008 MB003,TA007 MB004,'' MB005,TB013 MB007,TB009 MB008,TA003 MB009,''  MB010,('報價單'+TA001+'-'+TA002) MB012,(CASE WHEN TA022='1' THEN 'Y' ELSE 'N' END) MB013,'' MB014,0 MB015,0 MB016,TB016 MB017,'' MB018,'TK'  COMPANY ,TA005 CREATOR ,MV004 USR_GROUP ,CONVERT(NVARCHAR,GETDATE(),112) CREATE_DATE ,1 FLAG
                                        FROM [TK].dbo.COPTA,[TK].dbo.COPTB,[TK].dbo.CMSMV
                                        WHERE 1=1
                                        AND TA001=TB001 AND TA002=TB002
                                        AND MV001=TA005
                                        AND TA001=@TA001 AND TA002=@TA002
                                        AND TA004+TB004+TB008+TA007+TB016 NOT IN 
                                        (
                                        SELECT LTRIM(RTRIM(MB001))+LTRIM(RTRIM(MB002))+LTRIM(RTRIM(MB003))+LTRIM(RTRIM(MB004))+LTRIM(RTRIM(MB017))
                                        FROM [TK].dbo.COPMB
                                        )

                                        UPDATE [TK].dbo.COPMB
                                        SET MB009=@MB009,MB012=@MB012
                                        WHERE LTRIM(RTRIM(MB001))+LTRIM(RTRIM(MB002))+LTRIM(RTRIM(MB003))+LTRIM(RTRIM(MB004))+LTRIM(RTRIM(MB017)) IN 
                                        (
                                        SELECT TA004+TB004+TB008+TA007+TB016
                                        FROM [TK].dbo.COPTA,[TK].dbo.COPTB,[TK].dbo.CMSMV
                                        WHERE 1=1
                                        AND TA001=TB001 AND TA002=TB002
                                        AND MV001=TA005
                                        AND TA001=@TA001 AND TA002=@TA002
                                        )

                                        INSERT INTO [TK].dbo.COPMC
                                        (MC001, MC002, MC003, MC004, MC005, MC006, MC007, MC008, MC009,  COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG) 
                                        SELECT TA004 MC001,TB004 MC002,TB008 MC003,TA007 MC004,TK004 MC005,TK006 MC006,'' MC007,TK005 MC008,TB016 MC009,'TK'  COMPANY ,TA005 CREATOR ,MV004 USR_GROUP ,CONVERT(NVARCHAR,GETDATE(),112) CREATE_DATE ,1 FLAG
                                        FROM [TK].dbo.COPTA,[TK].dbo.COPTB,[TK].dbo.CMSMV,[TK].dbo.COPTK 
                                        WHERE 1=1
                                        AND TA001=TB001 AND TA002=TB002
                                        AND MV001=TA005
                                        AND TK001=TB001 AND TK002=TB002 AND TK003=TB003
                                        AND TA001=@TA001 AND TA002=@TA002
                                        AND TA004+TB004+TB008+TA007+TK004+TK006+TB016  NOT IN 
                                        (
                                        SELECT LTRIM(RTRIM(MC001))+LTRIM(RTRIM(MC002))+LTRIM(RTRIM(MC003))+LTRIM(RTRIM(MC004))+LTRIM(RTRIM(MC005))+LTRIM(RTRIM(MC009))
                                        FROM [TK].dbo.COPMC
                                        )

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
                    command.Parameters.Add("@TB011", SqlDbType.NVarChar).Value = "Y";
                    command.Parameters.Add("@TA016", SqlDbType.NVarChar).Value = "Y";
                    command.Parameters.Add("@TA019", SqlDbType.NVarChar).Value = "Y";
                    command.Parameters.Add("@TA029", SqlDbType.NVarChar).Value = "N";
                    command.Parameters.Add("@MB009", SqlDbType.NVarChar).Value = MB009;
                    command.Parameters.Add("@MB012", SqlDbType.NVarChar).Value = MB012;
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
