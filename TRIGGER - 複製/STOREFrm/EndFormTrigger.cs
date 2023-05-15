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

namespace TKUOF.TRIGGER.STOREFrm
{
    //門市導表

    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {

        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string FORMID = null;
            string MODIFIER = null;
            string TC001 = null;
           

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TC001 = applyTask.Task.CurrentDocument.Fields["TC001"].FieldValue.ToString().Trim();
                    FORMID = applyTask.FormNumber;
            MODIFIER = applyTask.Task.Applicant.Account;

            ///表單核準後  if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            ///單一節點核準後 if applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(TC001))
                {
                    ADDSTORESCHECK();
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void ADDSTORESCHECK()
        {
            string TC027 = "Y";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                  
                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TC027", SqlDbType.NVarChar).Value = TC027;
                   

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
