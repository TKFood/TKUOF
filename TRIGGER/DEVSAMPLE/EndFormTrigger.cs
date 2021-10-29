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
        public string DV02;
        public string DV03;
        public string DV04;
        public string DV05;
        public string DV06;
        public string DV07;
        public string DV08;
        public string DV09;
        public string DV10;
        public string DVV01;
        public string DVV02;
        public string DVV03;
        public string DVV04;
        public string DVV05;
        public string DVV06;
        public string DVV07;
        public string DVV08;
        public string ISCLOSE;

        public void Finally()
        {
            
        }

        public string GetFormResult(ApplyTask applyTask)
        {           
            FORMID = applyTask.FormNumber;
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            DV01 = applyTask.Task.CurrentDocument.Fields["DV01"].FieldValue.ToString().Trim();
            DV02 = applyTask.Task.CurrentDocument.Fields["DV02"].FieldValue.ToString().Trim();
            DV03 = applyTask.Task.CurrentDocument.Fields["DV03"].FieldValue.ToString().Trim();
            DV04 = "";
            DV05 = applyTask.Task.CurrentDocument.Fields["DV05"].FieldValue.ToString().Trim();
            DV06 = applyTask.Task.CurrentDocument.Fields["DV06"].FieldValue.ToString().Trim();
            DV07 = applyTask.Task.CurrentDocument.Fields["DV07"].FieldValue.ToString().Trim();
            DV08 = applyTask.Task.CurrentDocument.Fields["DV08"].FieldValue.ToString().Trim();
            DV09 = applyTask.Task.CurrentDocument.Fields["DV09"].FieldValue.ToString().Trim();
            DV10 = applyTask.Task.CurrentDocument.Fields["DV10"].FieldValue.ToString().Trim();
            ISCLOSE = "N";

            //DVV01 = applyTask.Task.CurrentDocument.Fields["DVV01"].FieldValue.ToString().Trim();



            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult== Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(FORMID) )
                {
                    foreach (XmlNode node in xmlDoc.SelectNodes("./Form/FormFieldValue/FieldItem[@fieldId='DETAILS']/DataGrid/Row"))
                    {
                        DVV01 = node.SelectSingleNode("./Cell[@fieldId='DVV01']").Attributes["fieldValue"].Value;
                        DVV02 = node.SelectSingleNode("./Cell[@fieldId='DVV02']").Attributes["fieldValue"].Value;
                        DVV03 = node.SelectSingleNode("./Cell[@fieldId='DVV03']").Attributes["fieldValue"].Value;
                        DVV04 = node.SelectSingleNode("./Cell[@fieldId='DVV04']").Attributes["fieldValue"].Value;
                        DVV05 = node.SelectSingleNode("./Cell[@fieldId='DVV05']").Attributes["fieldValue"].Value;
                        DVV06 = node.SelectSingleNode("./Cell[@fieldId='DVV06']").Attributes["fieldValue"].Value;
                        DVV07 = node.SelectSingleNode("./Cell[@fieldId='DVV07']").Attributes["fieldValue"].Value;
                        DVV08 = node.SelectSingleNode("./Cell[@fieldId='DVV08']").Attributes["fieldValue"].Value;

                        ADDTBSAMPLE(FORMID, DV01, DV02, DV03, DV04, DV05, DV06, DV07, DV08, DV09, DV10, DVV01, DVV02, DVV03, DVV04, DVV05, DVV06, DVV07, DVV08, ISCLOSE);
                    }

                           
                    
                }
            }
           

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public void ADDTBSAMPLE(string FORMID, string DV01, string DV02, string DV03, string DV04, string DV05, string DV06, string DV07, string DV08, string DV09, string DV10, string DVV01, string DVV02, string DVV03, string DVV04, string DVV05, string DVV06, string DVV07, string DVV08, string ISCLOSE)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                    INSERT INTO [TKRESEARCH].[dbo].[TBSAMPLE]
                                    ([FORMID],[DV01],[DV02],[DV03],[DV04],[DV05],[DV06],[DV07],[DV08],[DV09],[DV10]
                                    ,[DVV01],[DVV02],[DVV03],[DVV04],[DVV05],[DVV06],[DVV07],[DVV08]
                                    ,[ISCLOSE])
                                    VALUES 
                                    (@FORMID,@DV01,@DV02,@DV03,@DV04,@DV05,@DV06,@DV07,@DV08,@DV09,@DV10
                                    ,@DVV01,@DVV02,@DVV03,@DVV04,@DVV05,@DVV06,@DVV07,@DVV08
                                    ,@ISCLOSE)
                    
                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@FORMID", SqlDbType.NVarChar).Value = FORMID;
                    command.Parameters.Add("@DV01", SqlDbType.NVarChar).Value = DV01;
                    command.Parameters.Add("@DV02", SqlDbType.NVarChar).Value = DV02;
                    command.Parameters.Add("@DV03", SqlDbType.NVarChar).Value = DV03;
                    command.Parameters.Add("@DV04", SqlDbType.NVarChar).Value = DV04;
                    command.Parameters.Add("@DV05", SqlDbType.NVarChar).Value = DV05;
                    command.Parameters.Add("@DV06", SqlDbType.NVarChar).Value = DV06;
                    command.Parameters.Add("@DV07", SqlDbType.NVarChar).Value = DV07;
                    command.Parameters.Add("@DV08", SqlDbType.NVarChar).Value = DV08;
                    command.Parameters.Add("@DV09", SqlDbType.NVarChar).Value = DV09;
                    command.Parameters.Add("@DV10", SqlDbType.NVarChar).Value = DV10;
                    command.Parameters.Add("@DVV01", SqlDbType.NVarChar).Value = DVV01;
                    command.Parameters.Add("@DVV02", SqlDbType.NVarChar).Value = DVV02;
                    command.Parameters.Add("@DVV03", SqlDbType.NVarChar).Value = DVV03;
                    command.Parameters.Add("@DVV04", SqlDbType.NVarChar).Value = DVV04;
                    command.Parameters.Add("@DVV05", SqlDbType.NVarChar).Value = DVV05;
                    command.Parameters.Add("@DVV06", SqlDbType.NVarChar).Value = DVV06;
                    command.Parameters.Add("@DVV07", SqlDbType.NVarChar).Value = DVV07;
                    command.Parameters.Add("@DVV08", SqlDbType.NVarChar).Value = DVV08;
                    command.Parameters.Add("@ISCLOSE", SqlDbType.NVarChar).Value = ISCLOSE;

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
