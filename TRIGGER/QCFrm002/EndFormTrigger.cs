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

namespace TKUOF.TRIGGER.QCFrm002
{
    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {
            //throw new NotImplementedException();
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                ADDTB_WKF_EXTERNAL_TASK(applyTask);
            }
               

            return "";

            //throw new NotImplementedException();
        }

        public void OnError(Exception errorException)
        {
            //throw new NotImplementedException();
        }

        public void ADDTB_WKF_EXTERNAL_TASK(ApplyTask applyTask)
        {
            XmlDocument xmlDoc = new XmlDocument();
            //建立根節點
            XmlElement Form = xmlDoc.CreateElement("Form");
            Form.SetAttribute("formVersionId", "0dc91399-9488-4baf-94c5-28ede3d90206");
            Form.SetAttribute("urgentLevel", "2");
            //加入節點底下
            xmlDoc.AppendChild(Form);

            //建立節點
            XmlElement Applicant = xmlDoc.CreateElement("Applicant");
            Applicant.SetAttribute("account", applyTask.Task.Applicant.Account);
            Applicant.SetAttribute("groupId", applyTask.Task.Applicant.GroupId);
            Applicant.SetAttribute("jobTitleId", applyTask.Task.Applicant.JobTitleId);
            //加入節點底下
            Form.AppendChild(Applicant);

            //建立節點
            XmlElement FormFieldValue = xmlDoc.CreateElement("FormFieldValue");
            //加入至節點底下
            Form.AppendChild(FormFieldValue);

            //建立節點FieldItem
            XmlElement FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm001TxT");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);


            //建立DataGrid
            XmlElement DataGrid = xmlDoc.CreateElement("DataGrid");
            FormFieldValue.AppendChild(DataGrid);

            //建立DataGrid
            XmlElement Row = xmlDoc.CreateElement("Row");
            FormFieldValue.AppendChild(Row);
            Row.SetAttribute("order", "0");

            //建立節點Cell
            XmlElement Cell = xmlDoc.CreateElement("Cell");
            Cell = xmlDoc.CreateElement("Cell");
            Cell.SetAttribute("fieldId", "QC1002201000013" );
            Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002SN"].FieldValue.ToString().Trim());
            Cell.SetAttribute("realValue", "");
            Cell.SetAttribute("customValue", "");
            Cell.SetAttribute("enableSearch", "True");


            //Cell.SetAttribute("fieldId", "B02");
            //Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["A02"].FieldValue.ToString().Trim());
            //Cell.SetAttribute("realValue", "");
            //Cell.SetAttribute("enableSearch", "True");
            //Cell.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            //Cell.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            //Cell.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            //Cell.SetAttribute("fillSiteId", "");


            //加入至members節點底下
            FormFieldValue.AppendChild(Cell);

            //ADD TO DB
            string connectionString = ConfigurationManager.ConnectionStrings["connectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@" INSERT INTO [UOFTEST].dbo.TB_WKF_EXTERNAL_TASK");
            queryString.AppendFormat(@" (EXTERNAL_TASK_ID,FORM_INFO,STATUS)");
            queryString.AppendFormat(@" VALUES (NEWID(),@XML,2)");
            queryString.AppendFormat(@" ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@XML", SqlDbType.NVarChar).Value = Form.OuterXml;                    

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
