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

namespace TKUOF.TRIGGER.COPMA1003
{
    //訂單的核準

    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {
            
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string MA001= null;
            string MA033= null;
          

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            MA001 = applyTask.Task.CurrentDocument.Fields["sd004"].FieldValue.ToString().Trim();
            MA033 = applyTask.Task.CurrentDocument.Fields["sd010"].FieldValue.ToString().Trim();
  

            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult== Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(MA001) && !string.IsNullOrEmpty(MA033))
                {
                    UPDATECOPTMA(MA001, MA033);
                }
            }
           

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public void UPDATECOPTMA(string MA001, string MA033)
        {
        
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();



            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                    UPDATE [TK].dbo.COPMA
                                    SET MA033=@MA033, FLAG=FLAG+1,COMPANY=@COMPANY,MODIFIER=@MODIFIER ,MODI_DATE=@MODI_DATE, MODI_TIME=@MODI_TIME                                     
                                    WHERE MA001=@MA001

                                    ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@MA001", SqlDbType.NVarChar).Value = MA001;
                    command.Parameters.Add("@MA033", SqlDbType.NVarChar).Value = MA033;                    
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
