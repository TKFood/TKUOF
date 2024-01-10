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

namespace TKUOF.TRIGGER.PURTH01SIGNCHANGES
{
    //訂單的核準

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

            int rows = 1;
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
                        string  TD004 = cell.fieldValue;
                        EXESQL.AppendFormat(@" TD004='{0}'", TD004);
                    }
                    if (cell.fieldId == "INPCTS")
                    {
                        string INPCTS = cell.fieldValue;
                        EXESQL.AppendFormat(@" INPCTS='{0}'", INPCTS);
                    }

                }
            
                rows = rows + 1;

                EXESQL.AppendFormat(@" rows='{0}'", rows);
            }

            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(EXESQL.ToString()))
                {
                    UPDATE_INVMB_MB045(EXESQL.ToString());
                }
            }



            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATE_INVMB_MB045(string EXESQL)
        {
            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();



            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                   

                                        ");

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
