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

namespace TKUOF.TRIGGER.COPTCD
{
    //訂單的核準

    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {
            
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TC001=null;
            string TC002 = null;
            string FORMID= null;
            string MODIFIER = null;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TC001 = applyTask.Task.CurrentDocument.Fields["TC001"].FieldValue.ToString().Trim();
            TC002 = applyTask.Task.CurrentDocument.Fields["TC002"].FieldValue.ToString().Trim();
            FORMID = applyTask.FormNumber;
            MODIFIER = applyTask.Task.Applicant.Account;

            ///核準
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                if (!string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002))
                {
                    UPDATECOPTCD(TC001, TC002, FORMID, MODIFIER);
                }
            }
           

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public void UPDATECOPTCD(string TC001, string TC002, string FORMID, string MODIFIER)
        {
            string TC027 = "Y";
            string TC048 = "N";
            string TD021 = "Y";
            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                    UPDATE [TK].dbo.COPTC
                                    SET TC027=@TC027,TC048=@TC048, FLAG=FLAG+1,COMPANY=@COMPANY,MODIFIER=@MODIFIER ,MODI_DATE=@MODI_DATE, MODI_TIME=@MODI_TIME 
                                    ,UDF03=@FORMID
                                    WHERE TC001=@TC001 AND TC002=@TC002

                                    UPDATE [TK].dbo.COPTD 
                                    SET TD021=@TD021, FLAG=FLAG+1,COMPANY=@COMPANY,MODIFIER=@MODIFIER ,MODI_DATE=@MODI_DATE, MODI_TIME=@MODI_TIME 
                                    WHERE TD001=@TC001 AND TD002=@TC002
                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TA001", SqlDbType.NVarChar).Value = TC001;
                    command.Parameters.Add("@TA002", SqlDbType.NVarChar).Value = TC002;
                    command.Parameters.Add("@FORMID", SqlDbType.NVarChar).Value = FORMID;
                    command.Parameters.Add("@TC027", SqlDbType.NVarChar).Value = TC027;
                    command.Parameters.Add("@TC048", SqlDbType.NVarChar).Value = TC048;
                    command.Parameters.Add("@TD021", SqlDbType.NVarChar).Value = TD021;
                    command.Parameters.Add("@COMPANY", SqlDbType.NVarChar).Value = COMPANY;
                    command.Parameters.Add("@MODIFIER", SqlDbType.NVarChar).Value = MODIFIER;
                    command.Parameters.Add("@MODI_DATE", SqlDbType.NVarChar).Value = MODI_DATE;
                    command.Parameters.Add("@MODI_TIME", SqlDbType.NVarChar).Value = MODI_TIME;

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
