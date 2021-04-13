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
            //每次簽核時，簽準過後
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
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
            //<FieldItem fieldId="GMOFrm001BSC" enableSearch="True" fillerName="葉志剛" fillerUserGuid="0077c97a-8699-4688-be7e-ea1ecb960145" fillerAccount="120002" fillSiteId="ec110e59-c14e-45ac-8c16-a82060cfb697">
            //<Form formVersionId="77eb3223-97ce-4b5c-a536-1fa032561d4f">
            //<Applicant userGuid="0077c97a-8699-4688-be7e-ea1ecb960145" account="120002" name="葉志剛" />
            //  <FormFieldValue>
            //    <FieldItem fieldId="GMOFrm001BSC" enableSearch="True" fillerName="葉志剛" fillerUserGuid="0077c97a-8699-4688-be7e-ea1ecb960145" fillerAccount="120002" fillSiteId="ec110e59-c14e-45ac-8c16-a82060cfb697">
            //      <DataGrid>
            //        <Row order="0">
            //          <Cell fieldId="GMOFrm001BF1" fieldValue="40" realValue="" customValue="" enableSearch="True" />
            //          <Cell fieldId="GMOFrm001CB" fieldValue="20" realValue="" customValue="" enableSearch="True" />
            //          <Cell fieldId="GMOFrm001RG" fieldValue="15" realValue="" customValue="" enableSearch="True" />
            //          <Cell fieldId="GMOFrm001CR" fieldValue="15" realValue="" customValue="" enableSearch="True" />
            //          <Cell fieldId="GMOFrm001SF" fieldValue="10" realValue="" customValue="" enableSearch="True" />
            //          <Cell fieldId="GMOFrm001TO" fieldValue="100" realValue="" customValue="" enableSearch="True" />
            //          <Cell fieldId="GMOFrm001PU" fieldValue="林忠輝(120023)" realValue="&lt;UserSet&gt;&lt;Element type='user'&gt; &lt;userId&gt;858fd300-93c6-4a4b-a69a-c5ff125f71e4&lt;/userId&gt;&lt;/Element&gt;&lt;/UserSet&gt;&#xD;&#xA;" customValue="" enableSearch="True" />
            //        </Row>
            //      </DataGrid>
            //    </FieldItem>
            //  </FormFieldValue>
            //</Form>

            //針對不同的簽核站點做判斷
            //要在表單設計的簽核卡設定SiteCode
            if (applyTask.SiteCode != "X")
            { }

            //建立根節點
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);

            //計算Row有幾筆
            XmlElement xmlElem = xmlDoc.DocumentElement;//獲取根節點
            int rowscounts = xmlDoc.SelectNodes("./Form/FormFieldValue/FieldItem[@fieldId='GMOFrm001BSC']/DataGrid/Row").Count;

            //建立節點Row，把rowscounts+1
            //Row	
            XmlElement Row = xmlDoc.CreateElement("Row");
            Row.SetAttribute("order", (rowscounts).ToString());

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

            XmlNode DataGrid=xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='GMOFrm001BSC']/DataGrid");
            DataGrid.AppendChild(Row);


            //GMOFrm001SC1=null
            XmlElement FieldItem = (XmlElement)xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='GMOFrm001SC1']");          
            FieldItem.SetAttribute("fieldValue", null);
            //GMOFrm001SC2=null
            FieldItem = (XmlElement)xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='GMOFrm001SC2']");
            FieldItem.SetAttribute("fieldValue", null);
            //GMOFrm001SC3=null
            FieldItem = (XmlElement)xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='GMOFrm001SC3']");
            FieldItem.SetAttribute("fieldValue", null);
            //GMOFrm001SC4=null
            FieldItem = (XmlElement)xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='GMOFrm001SC4']");
            FieldItem.SetAttribute("fieldValue", null);
            //GMOFrm001SC5=null
            FieldItem = (XmlElement)xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='GMOFrm001SC5']");
            FieldItem.SetAttribute("fieldValue", null);
            //GMOFrm001SCU=null
            FieldItem = (XmlElement)xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='GMOFrm001SCU']");
            FieldItem.SetAttribute("fieldValue", null);


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
