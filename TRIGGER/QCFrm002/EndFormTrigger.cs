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
using Ede.Uof.Utility.Log;

namespace TKUOF.TRIGGER.QCFrm002
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
            //要記得改成正式-1 
            //要記得改成正式-2

            //要記得改成正式-1
            //測試ID = "80a70ec7-3fa6-4b6c-9d5a-2711a6563164";
            //正式ID="1cc71c35-0a2c-490c-b733-f887b7975b17"
            string ID = "80a70ec7-3fa6-4b6c-9d5a-2711a6563164";

            //將applyTask轉成xml再取值，只取到文字的部份，不包含字型
            XElement xe = XElement.Parse(applyTask.CurrentDocXML);

            XmlDocument xmlDoc = new XmlDocument();
            //建立根節點
            XmlElement Form = xmlDoc.CreateElement("Form");

            //formVersionId
            Form.SetAttribute("formVersionId", ID);
            //urgentLevel
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

            //用ADDTACK，直接啟動起單
            ADDTACK(Form);

            ////ADD TO DB
            //string connectionString = ConfigurationManager.ConnectionStrings["connectionstring"].ToString();

            //StringBuilder queryString = new StringBuilder();

            ////要記得改成正式-2
            ////UOFTEST
            ////
            //queryString.AppendFormat(@" INSERT INTO [UOFTEST].dbo.TB_WKF_EXTERNAL_TASK
            //                             (EXTERNAL_TASK_ID,FORM_INFO,STATUS)
            //                            VALUES (NEWID(),@XML,2)
            //                            ");
            //////UOF
            //////
            ////queryString.AppendFormat(@" INSERT INTO [UOF].dbo.TB_WKF_EXTERNAL_TASK
            ////                             (EXTERNAL_TASK_ID,FORM_INFO,STATUS)
            ////                            VALUES (NEWID(),@XML,2)
            ////                            ");

            //try
            //{
            //    using (SqlConnection connection = new SqlConnection(connectionString))
            //    {

            //        SqlCommand command = new SqlCommand(queryString.ToString(), connection);
            //        command.Parameters.Add("@XML", SqlDbType.NVarChar).Value = Form.OuterXml;

            //        command.Connection.Open();

            //        int count = command.ExecuteNonQuery();

            //        connection.Close();
            //        connection.Dispose();

            //    }
            //}
            //catch
            //{

            //}
            //finally
            //{

            //}


        }

        public void ADDTACK(XmlElement Form)
        {
            Ede.Uof.WKF.Utility.TaskUtilityUCO taskUCO = new Ede.Uof.WKF.Utility.TaskUtilityUCO();

            string result = taskUCO.WebService_CreateTask(Form.OuterXml);

            XElement resultXE = XElement.Parse(result);

            string status = "";
            string formNBR = "";
            string error = "";

            if (resultXE.Element("Status").Value == "1")
            {
                status = "起單成功!";
                formNBR = resultXE.Element("FormNumber").Value;

                Logger.Write("TEST", status + formNBR);

            }
            else
            {
                status = "起單失敗!";
                error = resultXE.Element("Exception").Element("Message").Value;

                Logger.Write("TEST", status + error + "\r\n" + Form.OuterXml);

                throw new Exception(status + error + "\r\n" + Form.OuterXml);

            }
        }

    //    public void ADD()
    //    {  //ADD TO DB
    //        string connectionString = ConfigurationManager.ConnectionStrings["connectionstring"].ToString();

    //        StringBuilder queryString = new StringBuilder();
    //        queryString.AppendFormat(@" INSERT INTO [UOF].dbo.TB_WKF_EXTERNAL_TASK
    //                                 (EXTERNAL_TASK_ID,FORM_INFO,STATUS)
    //                                VALUES (NEWID(),@XML,2)
    //                                ");

    //        try
    //        {
    //            using (SqlConnection connection = new SqlConnection(connectionString))
    //            {

    //                SqlCommand command = new SqlCommand(queryString.ToString(), connection);
    //                command.Parameters.Add("@XML", SqlDbType.NVarChar).Value = "";

    //                command.Connection.Open();

    //                int count = command.ExecuteNonQuery();

    //                connection.Close();
    //                connection.Dispose();

    //            }
    //        }
    //        catch
    //        {

    //        }
    //        finally
    //        {

    //        }
    //    }


    //    //public void ADDTB_WKF_EXTERNAL_TASK(ApplyTask applyTask)
    //    //{
    //    //    //將applyTask轉成xml再取值，只取到文字的部份，不包含字型
    //    //    XElement xe = XElement.Parse(applyTask.CurrentDocXML);

    //    //    XmlDocument xmlDoc = new XmlDocument();
    //    //    //建立根節點
    //    //    XmlElement Form = xmlDoc.CreateElement("Form");
    //    //    //測試的id
    //    //    Form.SetAttribute("formVersionId", "fd2485f3-f513-44cd-aa99-2f2d27ffbc92");
    //    //    //正式的id
    //    //    //Form.SetAttribute("formVersionId", "6a6dff6e-184a-4f34-846f-8dac406351cd");
    //    //    Form.SetAttribute("urgentLevel", "2");
    //    //    //加入節點底下
    //    //    xmlDoc.AppendChild(Form);

    //    //    ////建立節點Applicant
    //    //    XmlElement Applicant = xmlDoc.CreateElement("Applicant");
    //    //    Applicant.SetAttribute("account", applyTask.Task.Applicant.Account);
    //    //    Applicant.SetAttribute("groupId", applyTask.Task.Applicant.GroupId);
    //    //    Applicant.SetAttribute("jobTitleId", applyTask.Task.Applicant.JobTitleId);
    //    //    //加入節點底下
    //    //    Form.AppendChild(Applicant);

    //    //    //建立節點 Comment
    //    //    XmlElement Comment = xmlDoc.CreateElement("Comment");
    //    //    Comment.InnerText = "申請者意見";
    //    //    //加入至節點底下
    //    //    Applicant.AppendChild(Comment);

    //    //    //建立節點
    //    //    XmlElement FormFieldValue = xmlDoc.CreateElement("FormFieldValue");
    //    //    //加入至節點底下
    //    //    Form.AppendChild(FormFieldValue);

    //    //    //建立節點FieldItem
    //    //    XmlElement FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001TxT");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
    //    //    FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
    //    //    FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);


    //    //    //建立DataGrid
    //    //    //DataGrid是FieldItem的子節點
    //    //    XmlElement DataGrid = xmlDoc.CreateElement("DataGrid");
    //    //    FieldItem.AppendChild(DataGrid);

    //    //    //建立Row
    //    //    //Row應是DataGrid的子節點
    //    //    XmlElement Row = xmlDoc.CreateElement("Row");
    //    //    DataGrid.AppendChild(Row);
    //    //    Row.SetAttribute("order", "0");


    //    //    //建立節點Cell
    //    //    XmlElement Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm001SN");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002SN"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002Date");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Date"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002User");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002User"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002Dept");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Dept"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002Rank");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Rank"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002CUST");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002CUST"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002TEL");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002TEL"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002Add");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Add"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002RDate");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002RDate"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002PKG");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002PKG"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002CN");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002CN"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002PRD");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002PRD"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002MD");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002MD"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002ED");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002ED"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002OD");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002OD"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002BP");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002BP"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002Prove");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Prove"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002Range");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Range"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002RP");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002RP"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002Abns");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Abns"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);

    //    //    ////只取值，不含字型設定
    //    //    //var value = (from xl in xe.Elements("FormFieldValue").Elements("FieldItem")
    //    //    //             where xl.Attribute("fieldId").Value == "QCFrm002Abn"
    //    //    //             select xl
    //    //    //          ).FirstOrDefault().Attribute("fieldValue").Value;
    //    //    ////判斷value是否為NUll
    //    //    //string QCFrm002Abn = "";
    //    //    //if (!string.IsNullOrEmpty(value.ToString()))
    //    //    //{
    //    //    //    QCFrm002Abn = XElement.Parse(value).Value.ToString();
    //    //    //}

    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002Abn");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Abn"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);

    //    //    //value = (from xl in xe.Elements("FormFieldValue").Elements("FieldItem")
    //    //    //         where xl.Attribute("fieldId").Value == "QCFrm002Process"
    //    //    //         select xl
    //    //    //         ).FirstOrDefault().Attribute("fieldValue").Value;

    //    //    ////判斷value是否為NUll
    //    //    //string QCFrm002Process = "";
    //    //    //if (!string.IsNullOrEmpty(value.ToString()))
    //    //    //{
    //    //    //    QCFrm002Process = XElement.Parse(value).Value.ToString();
    //    //    //}

    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002Process");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Process"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);

    //    //    //value = (from xl in xe.Elements("FormFieldValue").Elements("FieldItem")
    //    //    //         where xl.Attribute("fieldId").Value == "QCFrm002PR"
    //    //    //         select xl
    //    //    //      ).FirstOrDefault().Attribute("fieldValue").Value;

    //    //    ////判斷value是否為NUll
    //    //    //string QCFrm002PR = "";
    //    //    //if (!string.IsNullOrEmpty(value.ToString()))
    //    //    //{
    //    //    //    QCFrm002PR = XElement.Parse(value).Value.ToString();
    //    //    //}
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002PR");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002PR"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002RD");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002RD"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002RV");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002RV"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002QCC");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002QCC"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);

    //    //    //value = (from xl in xe.Elements("FormFieldValue").Elements("FieldItem")
    //    //    //         where xl.Attribute("fieldId").Value == "QCFrm002RCA"
    //    //    //         select xl
    //    //    //       ).FirstOrDefault().Attribute("fieldValue").Value;
    //    //    ////判斷value是否為NUll
    //    //    //string QCFrm002RCA = "";
    //    //    //if (!string.IsNullOrEmpty(value.ToString()))
    //    //    //{
    //    //    //    QCFrm002RCA = XElement.Parse(value).Value.ToString();
    //    //    //}
    //    //    Cell = xmlDoc.CreateElement("Cell");
    //    //    Cell.SetAttribute("fieldId", "QCFrm002RCA");
    //    //    Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002RCA"].FieldValue.ToString().Trim());
    //    //    Cell.SetAttribute("realValue", "");
    //    //    Cell.SetAttribute("customValue", "");
    //    //    Cell.SetAttribute("enableSearch", "True");
    //    //    Row.AppendChild(Cell);

    //    //    ////Cell.SetAttribute("fieldId", "B02");
    //    //    ////Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["A02"].FieldValue.ToString().Trim());
    //    //    ////Cell.SetAttribute("realValue", "");
    //    //    ////Cell.SetAttribute("enableSearch", "True");
    //    //    ////Cell.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
    //    //    ////Cell.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
    //    //    ////Cell.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
    //    //    ////Cell.SetAttribute("fillSiteId", "");
    //    //    ////加入至members節點底下
    //    //    ////CELL應是Row的子節點
    //    //    ////Row.AppendChild(Cell);


    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001SN");
    //    //    FieldItem.SetAttribute("fieldValue", "");
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    FieldItem.SetAttribute("IsNeedAutoNbr", "false");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001Date");
    //    //    FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Date"].FieldValue.ToString().Trim());
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrmB001User");
    //    //    FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002User"].FieldValue.ToString().Trim());
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001Dept");
    //    //    FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Dept"].FieldValue.ToString().Trim());
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001Rank");
    //    //    FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Rank"].FieldValue.ToString().Trim());
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001CUST");
    //    //    FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002CUST"].FieldValue.ToString().Trim());
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001PNO");
    //    //    FieldItem.SetAttribute("fieldValue", "");
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001CN");
    //    //    FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002CN"].FieldValue.ToString().Trim());
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001PRD");
    //    //    FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002PRD"].FieldValue.ToString().Trim());
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001RDate");
    //    //    FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002RDate"].FieldValue.ToString().Trim());
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001MD");
    //    //    FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002MD"].FieldValue.ToString().Trim());
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001ND");
    //    //    FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002ED"].FieldValue.ToString().Trim());
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001Range");
    //    //    FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Range"].FieldValue.ToString().Trim());
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001HB");
    //    //    FieldItem.SetAttribute("fieldValue", "");
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001Abn");
    //    //    FieldItem.SetAttribute("fieldValue", "");
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001Process");
    //    //    FieldItem.SetAttribute("fieldValue", "");
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001RCA");
    //    //    FieldItem.SetAttribute("fieldValue", "");
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem 原因分析填寫人
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001RCAU");
    //    //    FieldItem.SetAttribute("fieldValue", applyTask.Task.Applicant.UserName.ToString());
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001SR");
    //    //    FieldItem.SetAttribute("fieldValue", "");
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001PA");
    //    //    FieldItem.SetAttribute("fieldValue", "");
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001PAU");
    //    //    FieldItem.SetAttribute("fieldValue", "");
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001Pred");
    //    //    FieldItem.SetAttribute("fieldValue", "");
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001Cmf2");
    //    //    FieldItem.SetAttribute("fieldValue", "");
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001Comp");
    //    //    FieldItem.SetAttribute("fieldValue", "");
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);
    //    //    //建立節點FieldItem
    //    //    FieldItem = xmlDoc.CreateElement("FieldItem");
    //    //    FieldItem.SetAttribute("fieldId", "QCFrm001Cmf1");
    //    //    FieldItem.SetAttribute("fieldValue", "");
    //    //    FieldItem.SetAttribute("realValue", "");
    //    //    FieldItem.SetAttribute("enableSearch", "True");
    //    //    FieldItem.SetAttribute("fillerName", "");
    //    //    FieldItem.SetAttribute("fillerUserGuid", "");
    //    //    FieldItem.SetAttribute("fillerAccount", "");
    //    //    FieldItem.SetAttribute("fillSiteId", "");
    //    //    //加入至members節點底下
    //    //    FormFieldValue.AppendChild(FieldItem);

    //    //    //ADD TO DB
    //    //    string connectionString = ConfigurationManager.ConnectionStrings["connectionstring"].ToString();

    //    //    StringBuilder queryString = new StringBuilder();
    //    //    queryString.AppendFormat(@" INSERT INTO [UOF].dbo.TB_WKF_EXTERNAL_TASK
    //    //                             (EXTERNAL_TASK_ID,FORM_INFO,STATUS)
    //    //                            VALUES (NEWID(),@XML,2)
    //    //                            ");

    //    //    try
    //    //    {
    //    //        using (SqlConnection connection = new SqlConnection(connectionString))
    //    //        {

    //    //            SqlCommand command = new SqlCommand(queryString.ToString(), connection);
    //    //            command.Parameters.Add("@XML", SqlDbType.NVarChar).Value = Form.OuterXml;                    

    //    //            command.Connection.Open();

    //    //            int count = command.ExecuteNonQuery();

    //    //            connection.Close();
    //    //            connection.Dispose();

    //    //        }
    //    //    }
    //    //    catch
    //    //    {

    //    //    }
    //    //    finally
    //    //    {

    //    //    }
    //    //}
    }
}
