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
using Ede.Uof.EIP.Organization.Util;
using Ede.Uof.EIP.Organization.Util;
using Ede.Uof.EIP.SystemInfo;

namespace TKUOF.TRIGGER.PURTH01BACKS
{
    //12.資材類表單，PURTH01.進貨超收單，結案後，超收率回=0.1=10%

    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {

        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string MB001 = null;
            string FORMID = null;
            string MODIFIER = null;
            StringBuilder EXESQL = new StringBuilder();

            UserUCO userUCO = new UserUCO();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            //MB001 = applyTask.Task.CurrentDocument.Fields["MB001"].FieldValue.ToString().Trim();
     
            FORMID = applyTask.FormNumber;
            //MODIFIER = applyTask.Task.Applicant.Account;

            //取得簽核人工號
            EBUser ebUser = userUCO.GetEBUser(Current.UserGUID);
            MODIFIER = ebUser.Account;

           
            string TD004 = null;
            string INPCTS = null;
            string INPCTS_NEW = null;
            //針對DETAIL抓出來的資料作處理
            //DataGrid的欄位名=DETAILS
            Ede.Uof.WKF.Design.FieldDataGrid grid = applyTask.Task.CurrentDocument.Fields["DETAILS"] as Ede.Uof.WKF.Design.FieldDataGrid;
            foreach (var row in grid.FieldDataGridValue.RowValueList)
            {           

                //在DataGrid裡，每個欄位=TD004、INPCTS 
                foreach (var cell in row.CellValueList)
                {
                    if (cell.fieldId == "TD004")
                    {
                        TD004 = cell.fieldValue;
                    }
                    if (cell.fieldId == "INPCTS")
                    {
                        INPCTS_NEW = "0.1";
                        //INPCTS = cell.fieldValue;

                        //decimal CAL_INPCTS=Convert.ToDecimal(INPCTS);
                        //decimal CAL_INPCTS_NEW=0;

                        //if (decimal.TryParse(INPCTS, out CAL_INPCTS))
                        //{
                        //    CAL_INPCTS_NEW = CAL_INPCTS / 100;
                        //    INPCTS_NEW = CAL_INPCTS_NEW.ToString();
                        //}
                        //else
                        //{
                        //    // 當INPCTS無法轉換為Decimal時的處理邏輯
                        //    // 可以選擇報錯、提供預設值等方式處理
                        //}
                    }
                }
            

                EXESQL.AppendFormat(@" 
                                    UPDATE [TK].dbo.INVMB
                                    SET MB045='{1}'
                                    WHERE MB001='{0}'

                                    ", TD004, INPCTS_NEW);
            }

            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                if (!string.IsNullOrEmpty(EXESQL.ToString()))
                {
                    UPDATE_INVMB_MB045(EXESQL);
                }
            }



            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATE_INVMB_MB045(StringBuilder EXESQL)
        {
            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
        
            queryString.AppendFormat(@"  
                                        ");

            queryString = EXESQL;


            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    //command.Parameters.Add("@MB001", SqlDbType.NVarChar).Value = MB001;   

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
