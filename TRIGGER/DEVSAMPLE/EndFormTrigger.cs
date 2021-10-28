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

//13.研發類表單
//1004.樣品製作申請單

namespace TKUOF.TRIGGER.DEVSAMPLE
{
    //訂單的核準

    class EndFormTrigger : ICallbackTriggerPlugin
    {
       

        public string FORMID;
        public string DV01;
        public string DVV01;

        public void Finally()
        {
            
        }

        public string GetFormResult(ApplyTask applyTask)
        {           
            FORMID = applyTask.FormNumber;
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            DV01 = applyTask.Task.CurrentDocument.Fields["DV01"].FieldValue.ToString().Trim();
            //DVV01 = applyTask.Task.CurrentDocument.Fields["DVV01"].FieldValue.ToString().Trim();

          

            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult== Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(FORMID) )
                {
                    foreach (XmlNode node in xmlDoc.SelectNodes("./Form/FormFieldValue/FieldItem[@fieldId='DETAILS']/DataGrid/Row"))
                    {
                        DVV01 = node.SelectSingleNode("./Cell[@fieldId='DVV01']").Attributes["fieldValue"].Value;

                        ADDTBSAMPLE(FORMID, DV01, DVV01);
                    }

                           
                    
                }
            }
           

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public void ADDTBSAMPLE(string FORMID, string DV01, string DVV01)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                    INSERT INTO [TKRESEARCH].[dbo].[TBSAMPLE]
                                    ([FORMID],[DV01],[DVV01])
                                    VALUES 
                                    (@FORMID,@DV01,@DVV01)
                    
                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@FORMID", SqlDbType.NVarChar).Value = FORMID;
                    command.Parameters.Add("@DV01", SqlDbType.NVarChar).Value = DV01;
                    command.Parameters.Add("@DVV01", SqlDbType.NVarChar).Value = DVV01;

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
