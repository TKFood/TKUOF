﻿using System;
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
using Ede.Uof.EIP.SystemInfo;

namespace TKUOF.TRIGGER.QCFrm002
{
    class EndFormTrigger : ICallbackTriggerPlugin
    {
        //要記得改成正式-1 
        //要記得改成正式-2
        //要記得改成正式-3
        //要記得改成正式-4

        //測試ID = "4e05b38e-fc76-43fa-846a-ac222f3f47f8"
        //正式ID = "0dc787cc-68a4-4af2-9c96-8dccc2397969"
        //測試DB DBNAME = "UOFTEST";
        //正式DB DBNAME = "UOF";

        //品保的1002單轉成10001

        //20211109 是在測試版，之後上正式前，要把ID、DBNAME改成正式的正式ID、UOF


        //正式的id
        //string ID = "0dc787cc-68a4-4af2-9c96-8dccc2397969";
        string ID = "";     
        string DBNAME = "UOF";

        //簽核人
        string account = Current.User.Account;
        string groupId = Current.User.GroupID;
        string jobTitleId = Current.User.JobTitleID;



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
                //ADD();
                //20210915 品保-1001品質異常單，需判斷1002客訴異常單的品保判定(QCFrmQCC)，若品保判斷異常成立才要追踨生產
                if (applyTask.Task.CurrentDocument.Fields["QCFrm002QCC"].FieldValue.ToString().Trim().Equals("成立"))
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

        public string SEARCHFORM_UOF_VERSION_ID(string FORM_NAME)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["connectionstring"].ToString();
            SqlConnection sqlConn = new SqlConnection(connectionString);
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            StringBuilder queryString = new StringBuilder();
            DataSet ds = new DataSet();

            try
            {
                //要記得改成正式-3
                queryString.AppendFormat(@" 
                                                
                                        SELECT TOP 1 RTRIM(LTRIM(TB_WKF_FORM_VERSION.FORM_VERSION_ID)) FORM_VERSION_ID,TB_WKF_FORM_VERSION.FORM_ID,TB_WKF_FORM_VERSION.VERSION,TB_WKF_FORM_VERSION.ISSUE_CTL
                                        ,TB_WKF_FORM.FORM_NAME
                                        FROM [UOF].dbo.TB_WKF_FORM_VERSION,[UOF].dbo.TB_WKF_FORM
                                        WHERE 1=1
                                        AND TB_WKF_FORM_VERSION.FORM_ID=TB_WKF_FORM.FORM_ID
                                        AND TB_WKF_FORM_VERSION.ISSUE_CTL=1
                                        AND FORM_NAME='{0}'
                                        ORDER BY TB_WKF_FORM_VERSION.FORM_ID,TB_WKF_FORM_VERSION.VERSION DESC

                                                ", FORM_NAME);

                adapter = new SqlDataAdapter(@"" + queryString, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    return ds.Tables["TEMPds1"].Rows[0]["FORM_VERSION_ID"].ToString();
                }
                else
                {
                    return "";
                }

            }
            catch
            {
                return "";
            }
            finally
            {

            }
        }

        //public string SEARCHFORM_VERSION_ID(string FORM_NAME)
        //{
        //    string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();
        //    SqlConnection sqlConn = new SqlConnection(connectionString);
        //    SqlDataAdapter adapter = new SqlDataAdapter();
        //    SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        //    StringBuilder queryString = new StringBuilder();
        //    DataSet ds = new DataSet();

        //    try
        //    {
        //        //要記得改成正式-3
        //        queryString.AppendFormat(@" 
        //                                    SELECT 
        //                                    RTRIM(LTRIM([FORM_VERSION_ID])) AS FORM_VERSION_ID
        //                                    ,[FORM_NAME]
        //                                    FROM [TKIT].[dbo].[UOF_FORM_VERSION_ID]
        //                                    WHERE [FORM_NAME]='{0}'
        //                                    ", FORM_NAME);

        //        adapter = new SqlDataAdapter(@"" + queryString, sqlConn);
        //        sqlCmdBuilder = new SqlCommandBuilder(adapter);
        //        sqlConn.Open();
        //        ds.Clear();
        //        adapter.Fill(ds, "TEMPds1");
        //        sqlConn.Close();


        //        if (ds.Tables["TEMPds1"].Rows.Count >= 1)
        //        {
        //            return ds.Tables["TEMPds1"].Rows[0]["FORM_VERSION_ID"].ToString();
        //        }
        //        else
        //        {
        //            return "";
        //        }

        //    }
        //    catch
        //    {
        //        return "";
        //    }
        //    finally
        //    {

        //    }
        //}

        public string SEARCHATTACH_ID(string OLDTASK_ID)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["connectionstring"].ToString();
            SqlConnection  sqlConn = new SqlConnection(connectionString);
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            StringBuilder queryString = new StringBuilder();
            DataSet ds = new DataSet();

            try
            {
                //要記得改成正式-3
                queryString.AppendFormat(@" 
                                        SELECT ISNULL(ATTACH_ID,'') ATTACH_ID
                                        FROM [{0}].[dbo].[TB_WKF_TASK]
                                        WHERE ISNULL(ATTACH_ID,'')<>'' AND TASK_ID='{1}'
                                        ", DBNAME, OLDTASK_ID);

                adapter = new SqlDataAdapter(@"" + queryString, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {

                    Ede.Uof.Utility.Log.Logger.Write("Task", "附件有找到TASK_ID:/ ATTACH_ID: "+OLDTASK_ID +"/"+ds.Tables["TEMPds1"].Rows[0]["ATTACH_ID"].ToString());

                    return ds.Tables["TEMPds1"].Rows[0]["ATTACH_ID"].ToString();                   
                }
                else
                {

                    Ede.Uof.Utility.Log.Logger.Write("Task", "沒有附件TASK_ID:"+ OLDTASK_ID);
                    return "";
                }
               
            }
            catch
            {
                Ede.Uof.Utility.Log.Logger.Write("找附件時有錯 TASK_ID:", OLDTASK_ID);
                return "";
            }
            finally
            {

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
            ID = SEARCHFORM_UOF_VERSION_ID("1001.客訴品質異常處理單");
            Form.SetAttribute("formVersionId", ID);
            //urgentLevel
            Form.SetAttribute("urgentLevel", "2");
                       

            //加入節點底下
            xmlDoc.AppendChild(Form);

            ////建立節點Applicant
            XmlElement Applicant = xmlDoc.CreateElement("Applicant");

            //表單建立者為原單申請人
            //Applicant.SetAttribute("account", applyTask.Task.Applicant.Account);
            //Applicant.SetAttribute("groupId", applyTask.Task.Applicant.GroupId);
            //Applicant.SetAttribute("jobTitleId", applyTask.Task.Applicant.JobTitleId);

            //表單建立者為原單最後核準人
            Applicant.SetAttribute("account", account);
            Applicant.SetAttribute("groupId", groupId);
            Applicant.SetAttribute("jobTitleId", jobTitleId);

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


            //建立節點FieldItem
            //QCFrm001ASN 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm001ASN");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.FormNumber);
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
            FieldItem.SetAttribute("fieldId", "QCFrm001User");
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
            FieldItem.SetAttribute("fieldId", "QCFrm002Cmf");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Cmf"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //原因分析 QCFrm002Abn > QCFrm002Abn
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002Abn");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Abn"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
            FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
            FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //原因 QCFrm002Abns
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "QCFrm002Abns");
            FieldItem.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Abns"].FieldValue.ToString().Trim());
            FieldItem.SetAttribute("realValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Abns"].RealValue.ToString().Trim());
            FieldItem.SetAttribute("customValue", applyTask.Task.CurrentDocument.Fields["QCFrm002Abns"].CustomValue.ToString().Trim());
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
            try
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
            catch
            {

            }
            finally
            {

            }
           
        }

    }

   
}
