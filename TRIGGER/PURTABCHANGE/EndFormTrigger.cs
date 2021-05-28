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

namespace TKUOF.TRIGGER.PURTABCHANGE
{
    //請購單變更的核準

    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {
            
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string FORMID = null;
            string TA001=null;
            string TA002 = null;
            string TA006 = null;
            string TB003 = null;
            string TB004 = null;
            string TB009 = null;
            string TB011 = null;
            string TB012 = null;
            string ADDSQL = null;




            ////找到明細的XML物件
            //XElement xe = XElement.Parse(applyTask.CurrentDocXML);
            //var nodes = (from xl in xe.Element("FormFieldValue").Elements("FieldItem")
            //             where xl.Attribute("fieldId").Value == "TB"
            //             select xl).FirstOrDefault();

            ////找出PO_DETAIL的所有ROW集合
            //foreach (var node in nodes.Element("DataGrid").Elements("Row"))
            //{
            //    //找出所有ROW的CELL
            //    var cells = node.Elements("Cell");

            //    var PO_ITEM01的值 = cells.Where(p => p.Attribute("fieldId").Value == "PO_ITEM01").FirstOrDefault().Attribute("fieldValue");
            //    var MATERIAL_011的值 = cells.Where(p => p.Attribute("fieldId").Value == "MATERIAL_01").FirstOrDefault().Attribute("fieldValue");
            //    //........以此類推
            //}

            ///核準
            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(applyTask.CurrentDocXML);
                FORMID = applyTask.FormNumber;
                TA001 = applyTask.Task.CurrentDocument.Fields["TA001"].FieldValue.ToString().Trim();
                TA002 = applyTask.Task.CurrentDocument.Fields["TA002"].FieldValue.ToString().Trim();
                TA006 = applyTask.Task.CurrentDocument.Fields["TA006"].FieldValue.ToString().Trim();


                //針對DETAIL抓出來的資料作處理
                Ede.Uof.WKF.Design.FieldDataGrid grid = applyTask.Task.CurrentDocument.Fields["TB"] as Ede.Uof.WKF.Design.FieldDataGrid;
                foreach (var row in grid.FieldDataGridValue.RowValueList)
                {
                    foreach (var cell in row.CellValueList)
                    {
                        if (cell.fieldId == "TB003")
                        {
                            TB003 = cell.fieldValue;
                        }
                        if (cell.fieldId == "TB004")
                        {
                            TB004 = cell.fieldValue;
                        }
                        if (cell.fieldId == "TB009")
                        {
                            TB009 = cell.fieldValue;
                        }
                        if (cell.fieldId == "TB011")
                        {
                            TB011 = cell.fieldValue;
                        }
                        if (cell.fieldId == "TB012")
                        {
                            TB012 = cell.fieldValue;
                        }
                    }

                    if (!string.IsNullOrEmpty(FORMID) && !string.IsNullOrEmpty(TA001) && !string.IsNullOrEmpty(TA002))
                    {
                        ADDSQL = ADDSQL+SETPURTABUOFCHANGE(FORMID, TA001, TA002, TA006, TB003, TB004, TB009, TB011, TB012);
                        ADDSQL = ADDSQL + " ";
                    }
                }

                ADDPURTABUOFCHANGE(FORMID,TA001,TA002, ADDSQL);
            }
       

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public string SETPURTABUOFCHANGE(string FORMID, string TA001, string TA002, string TA006, string TB003, string TB004, string TB009, string TB011, string TB012)
        {
            StringBuilder SQL = new StringBuilder();
            SQL.AppendFormat(@" 
                                INSERT INTO [TKPUR].[dbo].[PURTABUOFCHANGE]
                                ([FORMID],[TA001],[TA002],[TA006],[TB003],[TB004],[TB009],[TB011],[TB012])
                                VALUES
                                (@FORMID,'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')

                                ",  TA001, TA002, TA006, TB003, TB004, TB009, TB011, TB012);

            return SQL.ToString();
        }

        public void ADDPURTABUOFCHANGE(string FORMID,string TA001,string TA002,string ADDSQL)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            //StringBuilder queryString = new StringBuilder();
            //queryString.AppendFormat(@"
                                       
            //                            ", FORMID);

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(ADDSQL.ToString(), connection);
                    command.Parameters.Add("@FORMID", SqlDbType.NVarChar).Value = FORMID;
         
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
