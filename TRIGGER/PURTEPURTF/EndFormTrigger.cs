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

namespace TKUOF.TRIGGER.PURTEPURTF
{
    //請購單的核準

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

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TE001 = applyTask.Task.CurrentDocument.Fields["TE001"].FieldValue.ToString().Trim();
            TE002 = applyTask.Task.CurrentDocument.Fields["TE002"].FieldValue.ToString().Trim();
            TE003 = applyTask.Task.CurrentDocument.Fields["TE003"].FieldValue.ToString().Trim();
            FORMID = applyTask.FormNumber;

            ///核準
            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                if (!string.IsNullOrEmpty(TE001) && !string.IsNullOrEmpty(TE002) && !string.IsNullOrEmpty(TE003))
                {
                    UPDATEPURTEPURTF(TE001, TE002, TE003, FORMID);
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

        public void UPDATEPURTEPURTF(string TE001, string TE002, string TE003, string FORMID)
        {
            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");


            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                        --INSERT PURTD

                                        --UPDATE PURTD

                                        --UPDATE PURTC

                                        --更新PURTC的未稅、稅額、總金額、數量

                                        --如果變更單整理指定結案，原PURTC也指定結案

                                        --如果變更單單身指定結案，原PURTD也指定結案
                                      
                                        ", FORMID);

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TE001", SqlDbType.NVarChar).Value = TE001;
                    command.Parameters.Add("@TE002", SqlDbType.NVarChar).Value = TE002;
                    command.Parameters.Add("@TE003", SqlDbType.NVarChar).Value = TE003;

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
