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

namespace TKUOF.TRIGGER.PURTAB
{
    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {
            
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TA001=null;
            string TA002 = null;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TA001 = applyTask.Task.CurrentDocument.Fields["PURTATA001"].FieldValue.ToString().Trim();
            TA002 = applyTask.Task.CurrentDocument.Fields["PURTATA002"].FieldValue.ToString().Trim();

            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                if (!string.IsNullOrEmpty(TA001) && !string.IsNullOrEmpty(TA002))
                {
                    UPDATEPURTAB(TA001, TA002);
                }
            }
           
            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public void UPDATEPURTAB(string TA001,string TA002)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"UPDATE [TK].dbo.PURTA SET TA007='Y' WHERE TA001=@TA001 AND TA002=@TA002 ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TA001", SqlDbType.NVarChar).Value = TA001;
                    command.Parameters.Add("@TA002", SqlDbType.NVarChar).Value = TA002;

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
