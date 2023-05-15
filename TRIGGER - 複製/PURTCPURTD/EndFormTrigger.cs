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

namespace TKUOF.TRIGGER.PURTCPURTD
{
    //請購單的核準

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

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TC001 = applyTask.Task.CurrentDocument.Fields["TC001"].FieldValue.ToString().Trim();
            TC002 = applyTask.Task.CurrentDocument.Fields["TC002"].FieldValue.ToString().Trim();
            FORMID = applyTask.FormNumber;

            ///核準
            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                if (!string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002))
                {
                    UPDATEPURTCPURTD(TC001, TC002, FORMID);
                }
            }
            //作廢
            else if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Cancel)
            {
                if (!string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002))
                {
                    UPDATEPURTCPURTDCANCEL(TC001, TC002, FORMID);
                }
            }
            //否決
            else if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Reject)
            {
                if (!string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002))
                {
                    UPDATEPURTCPURTDREJECT(TC001, TC002, FORMID);
                }
            }
            //退簽
            else if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Return)
            {
                if (!string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002))
                {
                    UPDATEPURTCPURTDRETURN(TC001, TC002, FORMID);
                }
            }

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public void UPDATEPURTCPURTD(string TC001, string TC002, string FORMID)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                        UPDATE [TK].dbo.PURTC SET TC014='Y' WHERE TC001=@TC001 AND TC002=@TC002 
                                        UPDATE [TK].dbo.PURTD SET TD018='Y' WHERE TD001=@TC001 AND TD002=@TC002 

                                        UPDATE [TK].dbo.PURTC SET UDF02='{0}' WHERE TC001=@TC001 AND TC002=@TC002 
                                        ", FORMID);

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TC001", SqlDbType.NVarChar).Value = TC001;
                    command.Parameters.Add("@TC002", SqlDbType.NVarChar).Value = TC002;

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

        public void UPDATEPURTCPURTDCANCEL(string TC001, string TC002, string FORMID)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"                                       
                                       

                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TC001", SqlDbType.NVarChar).Value = TC001;
                    command.Parameters.Add("@TC002", SqlDbType.NVarChar).Value = TC002;

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

        public void UPDATEPURTCPURTDREJECT(string TC001, string TC002, string FORMID)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"                                      
                                       

                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TC001", SqlDbType.NVarChar).Value = TC001;
                    command.Parameters.Add("@TC002", SqlDbType.NVarChar).Value = TC002;

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

        public void UPDATEPURTCPURTDRETURN(string TC001, string TC002, string FORMID)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"                                      
                                        

                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TC001", SqlDbType.NVarChar).Value = TC001;
                    command.Parameters.Add("@TC002", SqlDbType.NVarChar).Value = TC002;

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
