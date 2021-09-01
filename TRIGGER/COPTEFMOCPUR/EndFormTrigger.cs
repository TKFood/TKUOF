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

namespace TKUOF.TRIGGER.COPTEFMOCPUR
{
    //訂單的核準

    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {
            
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TE001=null;
            string TE002 = null;
            string TE003 = null;
            string FORMID= null;
            string MODIFIER = null;
            string MOC = null;
            string PUR = null;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TE001 = applyTask.Task.CurrentDocument.Fields["TE001"].FieldValue.ToString().Trim();
            TE002 = applyTask.Task.CurrentDocument.Fields["TE002"].FieldValue.ToString().Trim();
            TE003 = applyTask.Task.CurrentDocument.Fields["TE003"].FieldValue.ToString().Trim();
            MOC = applyTask.Task.CurrentDocument.Fields["MOC"].FieldValue.ToString().Trim();
            PUR = applyTask.Task.CurrentDocument.Fields["PUR"].FieldValue.ToString().Trim();
            FORMID = applyTask.FormNumber;
            MODIFIER = applyTask.Task.Applicant.Account;

            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult== Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(TE001) && !string.IsNullOrEmpty(TE002) && !string.IsNullOrEmpty(TE003))
                {
                    if (!string.IsNullOrEmpty(MOC)||!string.IsNullOrEmpty(PUR)) 
                    {
                        UPDATECOPTEF(TE001, TE002, TE003, FORMID, MODIFIER, MOC, PUR);
                    }
                    
                }
            }
           

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public void UPDATECOPTEF(string TE001, string TE002, string TE003, string FORMID, string MODIFIER,string MOC,string PUR)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                    --更新變單表單的編號到COPTD、COPTE
                                    --更新PUR、MOC備註到COPTD、COPTE

                                    UPDATE [TK].dbo.COPTC
                                    SET UDF05=SUBSTRING((UDF05+' '+@MOC+' '+@PUR),1,250)
                                    WHERE TC001=@TC001 AND TC002=@TC002
 
                                    UPDATE [TK].dbo.COPTE
                                    SET UDF05=SUBSTRING((@MOC+' '+@PUR),1,250)
                                    WHERE TE001=@TE001 AND TE002=@TE002  AND TE003=@TE003
                    
                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TC001", SqlDbType.NVarChar).Value = TE001;
                    command.Parameters.Add("@TC002", SqlDbType.NVarChar).Value = TE002;
                    command.Parameters.Add("@TE001", SqlDbType.NVarChar).Value = TE001;
                    command.Parameters.Add("@TE002", SqlDbType.NVarChar).Value = TE002;
                    command.Parameters.Add("@TE003", SqlDbType.NVarChar).Value = TE003;                
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
