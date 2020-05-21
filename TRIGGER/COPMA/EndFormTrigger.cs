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
using System.Xml.Linq;

namespace TKUOF.TRIGGER.COPMA
{
    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {
            
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string MA001 = null;
            string MA002 = null;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
  
            //MA001 = applyTask.Task.CurrentDocument.Fields["MA001"].FieldValue.ToString().Trim();
            //MA002 = applyTask.Task.CurrentDocument.Fields["MA002"].FieldValue.ToString().Trim();

            XmlNode node= xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem/FieldValue");
    
            MA001 = node.Attributes["MA001"].Value;
            MA002 = node.Attributes["MA002"].Value;


            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                if (!string.IsNullOrEmpty(MA001) && !string.IsNullOrEmpty(MA002))
                {
                    ADDTKCOPMA(MA001, MA002);
                }
            }

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public void ADDTKCOPMA(string MA001, string MA002)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@" INSERT INTO [TK].dbo.COPMA");
            queryString.AppendFormat(@" (COMPANY,MA001,MA002)");
            queryString.AppendFormat(@" VALUES (@MA001,@MA001,@MA002)");
            queryString.AppendFormat(@" ");
            queryString.AppendFormat(@" ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@MA001", SqlDbType.NVarChar).Value = MA001;
                    command.Parameters.Add("@MA002", SqlDbType.NVarChar).Value = MA002;
                     

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
