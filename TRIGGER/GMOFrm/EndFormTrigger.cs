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
        //TKUOF.TRIGGER.GMOFrm.EndFormTrigger

        public void Finally()
        {
            //throw new NotImplementedException();
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                UPDATE_TB_WKF_TASK(applyTask);
            }


            return "";

            //throw new NotImplementedException();
        }

        public void OnError(Exception errorException)
        {
            //throw new NotImplementedException();
        }

        public void UPDATE_TB_WKF_TASK(ApplyTask applyTask)
        {

            //建立根節點
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);

            //判斷DataGrid是否存在
            //若不存在，就直接新增Row order="0" 的資料
            //若存在，要先算出Row有幾筆，再加1，新增Row

            XmlElement xmlElem = xmlDoc.DocumentElement;//獲取根節點
            XmlNodeList bodyNode = xmlElem.GetElementsByTagName("DataGrid");//取節點名bodyXmlNode

            //如果沒有DataGrid，要在節點GMOFrm001SCU後面，新增GMOFrm001BSC+DataGrid+ROWS+CELLS
            if (bodyNode.Count == 0)
            {
                //查詢子節點
                //GMOFrm001SCU
                XmlNode GMOFrm001SCU = xmlDoc.SelectSingleNode("GMOFrm001SCU");

                //建立節點FieldItem
                //GMOFrm001BSC	
                XmlElement FieldItem = xmlDoc.CreateElement("FieldItem");
                FieldItem.SetAttribute("fieldId", "GMOFrm001BSC");
                FieldItem.SetAttribute("fieldValue", "");
                FieldItem.SetAttribute("realValue", "");
                FieldItem.SetAttribute("enableSearch", "True");
                FieldItem.SetAttribute("fillerName", applyTask.Task.Applicant.UserName);
                FieldItem.SetAttribute("fillerUserGuid", applyTask.Task.Applicant.UserGUID);
                FieldItem.SetAttribute("fillerAccount", applyTask.Task.Applicant.Account);
                FieldItem.SetAttribute("fillSiteId", "");
                //加入至members節點底下
                GMOFrm001SCU.AppendChild(FieldItem);

                //建立節點DataGrid
                //DataGrid	
                XmlElement DataGrid = xmlDoc.CreateElement("DataGrid");
                //FieldItem
                FieldItem.AppendChild(DataGrid);

                //建立節點Row
                //Row	
                XmlElement Row = xmlDoc.CreateElement("Row");
                Row.SetAttribute("order", "0");
                //DataGrid
                DataGrid.AppendChild(Row);

                //建立節點Row
                //Row	GMOFrm001BF1
                XmlElement Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "GMOFrm001BF1");
                Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["GMOFrm001SC1"].FieldValue.ToString().Trim());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);
                //建立節點Row
                //Row	GMOFrm001CB
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "GMOFrm001CB");
                Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["GMOFrm001SC2"].FieldValue.ToString().Trim());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);
                //建立節點Row
                //Row	GMOFrm001RG
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "GMOFrm001RG");
                Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["GMOFrm001SC3"].FieldValue.ToString().Trim());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);
                //建立節點Row
                //Row	GMOFrm001CR
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "GMOFrm001CR");
                Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["GMOFrm001SC4"].FieldValue.ToString().Trim());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);
                //建立節點Row
                //Row	GMOFrm001SF
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "GMOFrm001SF");
                Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["GMOFrm001SC5"].FieldValue.ToString().Trim());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);
                //Row	GMOFrm001TO
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "GMOFrm001TO");
                Cell.SetAttribute("fieldValue", (Convert.ToInt32(applyTask.Task.CurrentDocument.Fields["GMOFrm001SC1"].FieldValue.ToString().Trim())+ Convert.ToInt32(applyTask.Task.CurrentDocument.Fields["GMOFrm001SC2"].FieldValue.ToString().Trim())+ Convert.ToInt32(applyTask.Task.CurrentDocument.Fields["GMOFrm001SC3"].FieldValue.ToString().Trim())+ Convert.ToInt32(applyTask.Task.CurrentDocument.Fields["GMOFrm001SC4"].FieldValue.ToString().Trim())+ Convert.ToInt32(applyTask.Task.CurrentDocument.Fields["GMOFrm001SC5"].FieldValue.ToString().Trim())).ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);
                //建立節點Row
                //Row	GMOFrm001PU
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "GMOFrm001PU");
                Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["GMOFrm001SCU"].FieldValue.ToString().Trim());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);
                //建立節點Row
       

            }
            //如果已有DataGrid，就直接新增ROWS+CELLS
            else
            {
                int COUNTS = 0;

                //找出Row有多少筆，再加1
                foreach (var ROWS in xmlElem.GetElementsByTagName("Row"))
                {
                    COUNTS++;
                }

                //查詢子節點
                //DataGrid
                XmlNode DataGrid = xmlDoc.SelectSingleNode("DataGrid");

                //建立節點Row
                //Row	
                XmlElement Row = xmlDoc.CreateElement("Row");
                Row.SetAttribute("order", COUNTS.ToString());
                //DataGrid
                DataGrid.AppendChild(Row);

                //建立節點Row
                //Row	GMOFrm001BF1
                XmlElement Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "GMOFrm001BF1");
                Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["GMOFrm001SC1"].FieldValue.ToString().Trim());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);
                //建立節點Row
                //Row	GMOFrm001CB
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "GMOFrm001CB");
                Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["GMOFrm001SC2"].FieldValue.ToString().Trim());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);
                //建立節點Row
                //Row	GMOFrm001RG
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "GMOFrm001RG");
                Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["GMOFrm001SC3"].FieldValue.ToString().Trim());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);
                //建立節點Row
                //Row	GMOFrm001CR
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "GMOFrm001CR");
                Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["GMOFrm001SC4"].FieldValue.ToString().Trim());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);
                //建立節點Row
                //Row	GMOFrm001SF
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "GMOFrm001SF");
                Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["GMOFrm001SC5"].FieldValue.ToString().Trim());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);
                //Row	GMOFrm001TO
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "GMOFrm001TO");
                Cell.SetAttribute("fieldValue", (Convert.ToInt32(applyTask.Task.CurrentDocument.Fields["GMOFrm001SC1"].FieldValue.ToString().Trim()) + Convert.ToInt32(applyTask.Task.CurrentDocument.Fields["GMOFrm001SC2"].FieldValue.ToString().Trim()) + Convert.ToInt32(applyTask.Task.CurrentDocument.Fields["GMOFrm001SC3"].FieldValue.ToString().Trim()) + Convert.ToInt32(applyTask.Task.CurrentDocument.Fields["GMOFrm001SC4"].FieldValue.ToString().Trim()) + Convert.ToInt32(applyTask.Task.CurrentDocument.Fields["GMOFrm001SC5"].FieldValue.ToString().Trim())).ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);
                //建立節點Row
                //Row	GMOFrm001PU
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "GMOFrm001PU");
                Cell.SetAttribute("fieldValue", applyTask.Task.CurrentDocument.Fields["GMOFrm001SCU"].FieldValue.ToString().Trim());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);
                //建立節點Row

            }

            //UPDATE_TB_WKF_TASK
            string connectionString = ConfigurationManager.ConnectionStrings["connectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();

            //UOFTEST
            queryString.AppendFormat(@" UPDATE [UOFTEST].dbo.TB_WKF_TASK
                                        SET CURRENT_DOC=@XML
                                        WHERE TASK_ID='4eee190c-8f45-476e-bb3b-581cfa81a470'
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
                    command.Parameters.Add("@XML", SqlDbType.NVarChar).Value = xmlDoc.OuterXml;

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
