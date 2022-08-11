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

namespace TKUOF.TRIGGER.PURTGPURTH
{
    //進貨單的核準

    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {
            
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TG001=null;
            string TG002 = null;
          
            string FORMID= null;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TG001 = applyTask.Task.CurrentDocument.Fields["TG001"].FieldValue.ToString().Trim();
            TG002 = applyTask.Task.CurrentDocument.Fields["TG002"].FieldValue.ToString().Trim();
           
            FORMID = applyTask.FormNumber;

            ///核準
            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                if (!string.IsNullOrEmpty(TG001) && !string.IsNullOrEmpty(TG002) )
                {
                    UPDATEPURTGPURTH(TG001, TG002, FORMID);
                }
            }
            //作廢
            else if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Cancel)
            {
                //if (!string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002))
                //{
                //    UPDATEPURTCPURTDCANCEL(TC001, TC002, FORMID);
                //}
            }
            //否決
            else if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Reject)
            {
            //    if (!string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002))
            //    {
            //        UPDATEPURTCPURTDREJECT(TC001, TC002, FORMID);
            //    }
            }
            //退簽
            else if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Return)
            {
                //if (!string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002))
                //{
                //    UPDATEPURTCPURTDRETURN(TC001, TC002, FORMID);
                //}
            }

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public void UPDATEPURTGPURTH(string TG001, string TG002, string FORMID)
        {
           

            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");


            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                        
                                      
                                        ", FORMID);

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);

                    command.Parameters.Add("@TG001", SqlDbType.NVarChar).Value = TG001;
                    command.Parameters.Add("@TG002", SqlDbType.NVarChar).Value = TG002;
            

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
