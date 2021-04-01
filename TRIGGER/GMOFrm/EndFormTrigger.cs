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
using System.Xml.Linq;

namespace TKUOF.TRIGGER.GMOFrm
{
    class EndFormTrigger : ICallbackTriggerPlugin
    {
        //TKUOF.TRIGGER.QCFrm002.EndFormTrigger

        public void Finally()
        {
            //throw new NotImplementedException();
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                //ADD();
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
            //將applyTask轉成xml再取值，只取到文字的部份，不包含字型
            XElement xe = XElement.Parse(applyTask.CurrentDocXML);

            XmlDocument xmlDoc = new XmlDocument();
            //建立根節點
            XmlElement Form = xmlDoc.CreateElement("Form");
            //測試的id
            Form.SetAttribute("formVersionId", "56ea12d6-ba6b-4e2b-8ae0-9e82f47298e0");
            //正式的id
            //Form.SetAttribute("formVersionId", "6a6dff6e-184a-4f34-846f-8dac406351cd");
            Form.SetAttribute("urgentLevel", "2");
            //加入節點底下
            xmlDoc.AppendChild(Form);

            ////建立節點Applicant
            XmlElement Applicant = xmlDoc.CreateElement("Applicant");
            Applicant.SetAttribute("account", applyTask.Task.Applicant.Account);
            Applicant.SetAttribute("groupId", applyTask.Task.Applicant.GroupId);
            Applicant.SetAttribute("jobTitleId", applyTask.Task.Applicant.JobTitleId);
            //加入節點底下
            Form.AppendChild(Applicant);

            //建立節點 Comment
            XmlElement Comment = xmlDoc.CreateElement("Comment");
            Comment.InnerText = "申請者意見";
            //加入至節點底下
            Applicant.AppendChild(Comment);

            //建立節點
            XmlElement FormFieldValue = xmlDoc.CreateElement("FormFieldValue");
            //加入至節點底下
            Form.AppendChild(FormFieldValue);

            //建立節點FieldItem
            //QCFrm001SN 表單編號	
            XmlElement FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm001SN");
            FieldItem.SetAttribute("fieldValue", "");
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");

            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //日期QCFrm002Date>QCFrm001Date 
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm001Date");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Date"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");

            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //申請者 QCFrm002User> QCFrmB001User
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrmB001User");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002User"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //部門 QCFrm002Dept> QCFrm001Dept
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm001Dept");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Dept"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //職級 QCFrm002Rank> QCFrm001Rank
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm001Rank");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Rank"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //供應商/部門單位 QCFrm002CU> QCFrm001CUST
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm001CUST");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002CU"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //批號 QCFrm002PNO> QCFrm001PNO
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm001PNO");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002PNO"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //品號 QCFrm002CN> QCFrm001CN
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm001CN");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002CN"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //品名 QCFrm002PRD> QCFrm001PRD
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm001PRD");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002PRD"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //受理日期 QCFrm002RDate> QCFrm001RDate
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm001RDate");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002RDate"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //製造日期 QCFrm002MD> QCFrm001MD
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm001MD");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002MD"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //有效日期 QCFrm002ED> QCFrm001ND
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm001ND");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002ED"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //原因分析 QCFrm002Cmf> QCFrm001RCA
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm001RCA");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Cmf"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //原因分析填寫人 QCFrm002RCAU > QCFrm001RCAU
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm001RCAU");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002RCAU"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //ADD TO DB
            string connectionString = ConfigurationManager.ConnectionStrings["connectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();

            //UOFTEST
            queryString.AppendFormat(@" INSERT INTO [UOFTEST].dbo.TB_WKF_EXTERNAL_TASK
                                         (EXTERNAL_TASK_ID,FORM_INFO,STATUS)
                                        VALUES (NEWID(),@XML,2)
                                        ");
            //UOF 
            //queryString.AppendFormat(@" INSERT INTO [UOF].dbo.TB_WKF_EXTERNAL_TASK
            //                             (EXTERNAL_TASK_ID,FORM_INFO,STATUS)
            //                            VALUES (NEWID(),@XML,2)
            //                            ");

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
