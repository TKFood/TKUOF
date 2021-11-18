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
using Ede.Uof.Utility.FileCenter;
using Ede.Uof.Utility.FileCenter.V3;

namespace TKUOF.TRIGGER.QCFrm003
{
    class EndFormTrigger : ICallbackTriggerPlugin
    {
        //品保的1003單轉成1002


        //要記得改成正式-1 
        //要記得改成正式-2
        //要記得改成正式-3
        //要記得改成正式-4

        //測試ID = "";
        //正式ID = ""
        //測試DB DBNAME = "UOFTEST";
        //正式DB DBNAME = "UOF";



        //20211109 是在測試版，之後上正式前，要把ID、DBNAME改成正式的正式ID、UOF

        string ID = "860582f0-2a5e-44c2-8420-a4e0dbe397fc";
        string DBNAME = "UOFTEST";

        //TKUOF.TRIGGER.QCFrm002.EndFormTrigger

        string OLDTASK_ID = null;
        string NEWTASK_ID = null;
        string ATTACH_ID = null;

        public void Finally()
        {
            //throw new NotImplementedException();
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {



                //ADDTB_WKF_EXTERNAL_TASK(applyTask);
                //20210915 品保-1001品質異常單，需判斷1002客訴異常單的品保判定(QCFrmQCC)，若品保判斷異常成立才要追踨生產
                if (applyTask.Task.CurrentDocument.Fields["QCFrm003CC"].FieldValue.ToString().Trim().Equals("異常成立-品質異常相關。轉1002.客訴異常處理單 追蹤"))
                {

                    OLDTASK_ID = applyTask.TaskId;
                    ATTACH_ID = SEARCHATTACH_ID(OLDTASK_ID);

                    ADDTB_WKF_EXTERNAL_TASK(applyTask);

                    if (!string.IsNullOrEmpty(NEWTASK_ID) && !string.IsNullOrEmpty(ATTACH_ID))
                    {
                        ADDATTACH_IDTONEWTASK_ID(NEWTASK_ID, ATTACH_ID);
                    }
                }

            }


            return "";

            //throw new NotImplementedException();
        }

        public string SEARCHATTACH_ID(string OLDTASK_ID)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["connectionstring"].ToString();
            SqlConnection sqlConn = new SqlConnection(connectionString);
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            StringBuilder queryString = new StringBuilder();
            DataSet ds = new DataSet();

            //要記得改成正式-3
            queryString.AppendFormat(@" 
                                    SELECT ATTACH_ID
                                    FROM [{0}].[dbo].[TB_WKF_TASK]
                                    WHERE TASK_ID='{1}'
                                    ", DBNAME, OLDTASK_ID);

            adapter = new SqlDataAdapter(@"" + queryString, sqlConn);
            sqlCmdBuilder = new SqlCommandBuilder(adapter);
            sqlConn.Open();
            ds.Clear();
            adapter.Fill(ds, "TEMPds1");
            sqlConn.Close();


            if (ds.Tables["TEMPds1"].Rows.Count >= 0)
            {
                return ds.Tables["TEMPds1"].Rows[0]["ATTACH_ID"].ToString();
            }
            else
            {
                return null;
            }

        }


        public void OnError(Exception errorException)
        {
            //throw new NotImplementedException();
        }

        public void ADDTB_WKF_EXTERNAL_TASK(ApplyTask applyTask)
        {


            //要記得改成正式-1


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
            //QCFrm002SN 表單編號	
            XmlElement FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002SN");
            FieldItem.SetAttribute("fieldValue", "");
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");

            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);


            //建立節點FieldItem
            //QCFrm002SFNO 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002SFNO");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.FormNumber);
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");

            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003Date > QCFrm002Date 
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002Date");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003Date"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003User >  QCFrm002User
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002User");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003User"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003Dept >  QCFrm002Dept
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002Dept");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003Dept"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003Rank >  QCFrm002Rank
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002Rank");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003Rank"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003CUST >  QCFrm002CUST
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002CUST");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003CUST"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003TEL >  QCFrm002TEL
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002TEL");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003TEL"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003Add >  QCFrm002Add
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002Add");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003Add"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003CU >  QCFrm002CU
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002CU");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003CU"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003PNO >  QCFrm002PNO
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002PNO");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003PNO"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003CN >  QCFrm002CN
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002CN");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003CN"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003RDate >  QCFrm002RDate
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002RDate");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003RDate"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003PRD >  QCFrm002PRD
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002PRD");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003PRD"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003PKG >  QCFrm002PKG
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002PKG");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003PKG"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003MD >  QCFrm002MD
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002MD");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003MD"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003ED >  QCFrm002ED
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002ED");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003ED"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003OD >  QCFrm002OD
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002OD");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003OD"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003BP >  QCFrm002BP
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002BP");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003BP"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003Prove >  QCFrm002Prove
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002Prove");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003Prove"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003Abns >  QCFrm002Abns
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002Abns");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003Abns"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003Range >  QCFrm002Range
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002Range");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003Range"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003RP >  QCFrm002RP
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002RP");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003RP"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003RD >  QCFrm002RD
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002RD");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003RD"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003Abn >  QCFrm002Abn
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002Abn");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003Abn"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //QCFrm003Process >  QCFrm002Process
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002Process");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm003Process"].FieldValue.ToString().Trim());
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
                NEWTASK_ID = formNBR;

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


        public void ADDATTACH_IDTONEWTASK_ID(string NEWTASK_ID, string ATTACH_ID)
        {
            var newID = FileCenter.Clone(ATTACH_ID, Module.WKF);

            string connectionString = ConfigurationManager.ConnectionStrings["connectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();

            //要記得改成正式-4
            queryString.AppendFormat(@"
                                        UPDATE [{0}].dbo.TB_WKF_TASK
                                        SET ATTACH_ID='{2}'
                                        WHERE DOC_NBR='{1}'
                                        ", DBNAME, NEWTASK_ID, newID.ToString());

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);

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
