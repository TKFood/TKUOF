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

namespace TKUOF.TRIGGER.QCPURTGPURTH
{
    //品保檢驗進貨單明細

    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {
            
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            DataTable dt = new DataTable("QCHECK");
            dt.Columns.Add(new DataColumn("TH003", typeof(string)));
            dt.Columns.Add(new DataColumn("TH015", typeof(string)));
            dt.Columns.Add(new DataColumn("CHECK", typeof(string)));

            string TG001=null;
            string TG002 = null;
            string FORMID= null;
            string QCMAN = null;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TG001 = applyTask.Task.CurrentDocument.Fields["TG001"].FieldValue.ToString().Trim();
            TG002 = applyTask.Task.CurrentDocument.Fields["TG002"].FieldValue.ToString().Trim();
            FORMID = applyTask.FormNumber;
            QCMAN= applyTask.Task.Applicant.UserName;

            //品保人員簽核就啟動，不用等整張表單簽完
            //核準
            //記錄TH003,TH015,CHECK
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (applyTask.SiteCode.Equals("QCCHECK"))
                {
                    if (!string.IsNullOrEmpty(TG001) && !string.IsNullOrEmpty(TG002))
                    {
                        foreach (XmlNode node in xmlDoc.SelectNodes("./Form/FormFieldValue/FieldItem[@fieldId='PURTH']/DataGrid/Row"))
                        {
                            string TH003 = node.SelectSingleNode("./Cell[@fieldId='TH003']").Attributes["fieldValue"].Value;
                            string TH015 = node.SelectSingleNode("./Cell[@fieldId='TH015']").Attributes["fieldValue"].Value;
                            string CHECK = node.SelectSingleNode("./Cell[@fieldId='CHECK']").Attributes["fieldValue"].Value;

                            DataRow dr = dt.NewRow();
                            dr["TH003"] = TH003;
                            dr["TH015"] = TH015;
                            dr["CHECK"] = CHECK;
                            dt.Rows.Add(dr);
                        }

                        if (dt.Rows.Count > 0)
                        {
                            UPDATEPURTGPURTH(TG001, TG002, FORMID, QCMAN, dt);
                        }
                    }
                    

                }

            }
            //作廢
            else if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Cancel)
            {
              
            }
            //否決
            else if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Reject)
            {
               
            }
            //退簽
            else if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Return)
            {
              
            }

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public void UPDATEPURTGPURTH(string TG001, string TG002, string FORMID,string QCMAN,DataTable dt)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();

            if(dt.Rows.Count>0)
            {
                foreach(DataRow dr in dt.Rows)
                {
                    queryString.AppendFormat(@"
                                            UPDATE [TK].dbo.PURTH
                                            SET TH015='{1}',UDF01='{2}'
                                            WHERE TH001=@TG001 AND TH002=@TG002 AND TH003='{0}'
                                         ",  dr["TH003"].ToString(), dr["TH015"].ToString(), dr["CHECK"].ToString()+','+ QCMAN+'-'+ FORMID);
                }
               
            }


            //queryString.AppendFormat(@"
                                      
            //                            ", FORMID);

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
