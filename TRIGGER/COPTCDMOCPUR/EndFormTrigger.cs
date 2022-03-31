using System;
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

namespace TKUOF.TRIGGER.COPTCDMOCPUR
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
            string MOC = null;
            string PUR = null;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TC001 = applyTask.Task.CurrentDocument.Fields["TC001"].FieldValue.ToString().Trim();
            TC002 = applyTask.Task.CurrentDocument.Fields["TC002"].FieldValue.ToString().Trim();
            MOC = applyTask.Task.CurrentDocument.Fields["MOC"].FieldValue.ToString().Trim();
            PUR = applyTask.Task.CurrentDocument.Fields["PUR"].FieldValue.ToString().Trim();
            FORMID = applyTask.FormNumber;
            MODIFIER = applyTask.Task.Applicant.Account;

            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult== Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002))
                {
                    //MOC或PUR，其中一欄不是空的，就將MOC+PUR回寫到ERP訂單
                    if (!string.IsNullOrEmpty(MOC)||!string.IsNullOrEmpty(PUR))
                    {
                        UPDATECOPTCD(TC001, TC002, FORMID, MODIFIER, MOC, PUR);
                    }
                   
                }
            }
           

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public void UPDATECOPTCD(string TC001, string TC002, string FORMID, string MODIFIER,string MOC,string PUR)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();


            if (!string.IsNullOrEmpty(MOC))
            {
                MOC = DateTime.Now.ToString("MM/dd") + ":" + MOC + " ";
            }
            if (!string.IsNullOrEmpty(PUR))
            {
                PUR = DateTime.Now.ToString("MM/dd") + ":" + PUR + " ";
            }

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                    UPDATE [TK].dbo.COPTC
                                    SET UDF05=SUBSTRING((UDF05+' '+@MOC+' '+@PUR+' '),1,250)
                                    WHERE TC001=@TC001 AND TC002=@TC002

                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TC001", SqlDbType.NVarChar).Value = TC001;
                    command.Parameters.Add("@TC002", SqlDbType.NVarChar).Value = TC002;                    
                    command.Parameters.Add("@MOC", SqlDbType.NVarChar).Value = MOC;
                    command.Parameters.Add("@PUR", SqlDbType.NVarChar).Value = PUR;

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
