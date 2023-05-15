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

namespace TKUOF.TRIGGER.MOCI02
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
            string FORMID = null;
            string MODIFIER = null;
            UserUCO userUCO = new UserUCO();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TA001 = applyTask.Task.CurrentDocument.Fields["TA001"].FieldValue.ToString().Trim();
            TA002 = applyTask.Task.CurrentDocument.Fields["TA002"].FieldValue.ToString().Trim();
            FORMID = applyTask.FormNumber;
            //MODIFIER = applyTask.Task.Applicant.Account;

            //取得簽核人工號
            EBUser ebUser = userUCO.GetEBUser(Current.UserGUID);
            MODIFIER = ebUser.Account;
                       

            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(TA001) && !string.IsNullOrEmpty(TA002))
                {
                    UPDATEMOCTAMOCTB(TA001, TA002, FORMID, MODIFIER);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATEMOCTAMOCTB(string TA001,string TA002, string FORMID, string MODIFIER)
        {
            string TA013 = "Y";
            string TC048 = "N";
            string TB018 = "Y";
            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();



            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"                                   
                                        UPDATE [TK].dbo.MOCTA
                                        SET TA013=@TA013,TA040=TA003,TA041=@TA041,TA049=@TA049
                                        ,UDF03=@UDF03
                                        ,FLAG=FLAG+1
                                        ,COMPANY=@COMPANY,MODIFIER=@MODIFIER ,MODI_DATE=@MODI_DATE, MODI_TIME=@MODI_TIME
                                        WHERE TA001=@TA001 AND TA002=@TA002

                                        UPDATE [TK].dbo.MOCTB
                                        SET TB018=@TB018
                                        ,FLAG=FLAG+1
                                        ,COMPANY=@COMPANY,MODIFIER=@MODIFIER ,MODI_DATE=@MODI_DATE, MODI_TIME=@MODI_TIME 
                                        WHERE TB001=@TB001 AND TB002=@TB002

 
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
                    command.Parameters.Add("@TA013", SqlDbType.NVarChar).Value = TA013;
                   
                    command.Parameters.Add("@TA041", SqlDbType.NVarChar).Value = MODIFIER;
                    command.Parameters.Add("@TA049", SqlDbType.NVarChar).Value = "N";
                    command.Parameters.Add("@TB018", SqlDbType.NVarChar).Value = TB018;

                    command.Parameters.Add("@COMPANY", SqlDbType.NVarChar).Value = "TK";
                    command.Parameters.Add("@MODIFIER", SqlDbType.NVarChar).Value = MODIFIER;
                    command.Parameters.Add("@MODI_DATE", SqlDbType.NVarChar).Value = DateTime.Now.ToString("yyyyMMdd");
                    command.Parameters.Add("@MODI_TIME", SqlDbType.NVarChar).Value = DateTime.Now.ToString("HH:mm:ss");
                    command.Parameters.Add("@UDF03", SqlDbType.NVarChar).Value =FORMID;



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
