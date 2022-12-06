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

namespace TKUOF.TRIGGER.QCINVTAINVTB
{
    //品保檢驗客供明細

    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {
            
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            DataTable dt = new DataTable("QCHECK");
            dt.Columns.Add(new DataColumn("TB003", typeof(string)));
            dt.Columns.Add(new DataColumn("TH015", typeof(string)));
            dt.Columns.Add(new DataColumn("CHECK", typeof(string)));

            string TA001=null;
            string TA002 = null;
            string FORMID= null;
            string QCMAN = null;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TA001 = applyTask.Task.CurrentDocument.Fields["TA001"].FieldValue.ToString().Trim();
            TA002 = applyTask.Task.CurrentDocument.Fields["TA002"].FieldValue.ToString().Trim();
            FORMID = applyTask.FormNumber;
            //QCMAN= applyTask.Task.Applicant.UserName;

            //取得簽核人員
            //QCMAN = applyTask.Task.CurrentSigner.UserName;

            //品保人員簽核就啟動，不用等整張表單簽完
            //核準
            //記錄TH003,TH015,CHECK
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve )
            {
                if (applyTask.SiteCode.Equals("QCCHECK"))
                {
                    if (!string.IsNullOrEmpty(TA001) && !string.IsNullOrEmpty(TA002))
                    {
                        foreach (XmlNode node in xmlDoc.SelectNodes("./Form/FormFieldValue/FieldItem[@fieldId='INVTB']/DataGrid/Row"))
                        {
                            string TB003 = node.SelectSingleNode("./Cell[@fieldId='TB003']").Attributes["fieldValue"].Value;
                            string TH015 = node.SelectSingleNode("./Cell[@fieldId='TH015']").Attributes["fieldValue"].Value;
                            string CHECK = node.SelectSingleNode("./Cell[@fieldId='CHECK']").Attributes["fieldValue"].Value;

                            //取得簽核人員
                            QCMAN = applyTask.Task.CurrentSigner.UserName;

                            DataRow dr = dt.NewRow();
                            dr["TB003"] = TB003;
                            dr["TH015"] = TH015;
                            dr["CHECK"] = CHECK;
                            dt.Rows.Add(dr);
                        }

                        if (dt.Rows.Count > 0)
                        {
                            UPDATEINVTAINVTB(TA001, TA002, FORMID, QCMAN, dt);                          
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


        /// <summary>
        /// 更新進貨單的單身驗收、驗退數量、驗收狀況、驗收人員
        /// </summary>
        /// <param name="TG001"></param>
        /// <param name="TG002"></param>
        /// <param name="FORMID"></param>
        /// <param name="QCMAN"></param>
        /// <param name="dt"></param>
        public void UPDATEINVTAINVTB(string TA001, string TA002, string FORMID,string QCMAN,DataTable dt)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();

            if(dt.Rows.Count>0)
            {
                foreach(DataRow dr in dt.Rows)
                {

                    //設定TH028 檢驗狀態
                    //TH028='1' 待驗
                    //TH028='2' 合格
                    //TH028='3' 不良
                    string TH028 = "1";

                    if (dr["CHECK"].ToString().Equals("Y"))
                    {
                        TH028 = "2";
                    }
                    else
                    {
                        TH028 = "3";
                    }
                    queryString.AppendFormat(@"
                                            UPDATE [TK].dbo.INVTB
                                            SET TB007={1},UDF01='{2}'
                                            WHERE TB001=@TA001 AND TB002=@TA002 AND TB003='{0}'
                                         ",  dr["TB003"].ToString(), Convert.ToDecimal(dr["TH015"].ToString()), dr["CHECK"].ToString()+','+ QCMAN+'-'+ FORMID, Convert.ToDecimal(dr["TH015"].ToString()));
                }
               
            }


            //queryString.AppendFormat(@"
                                      
            //                            ", FORMID);

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
