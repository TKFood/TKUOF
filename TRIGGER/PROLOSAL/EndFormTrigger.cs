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

namespace TKUOF.TRIGGER.PROLOSAL
{
    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public class DATAPROLOSAL
        {
            /// <summary>
            /// TaskId=表單id
            /// GMOFrm001AVG6=平均分數
            /// GMOFrm001SB=總經理分數
            /// GMOFrm001BS=等級
            /// </summary>
            ///  
            public string TaskId;
            public string GMOFrm001AVG6;
            public string GMOFrm001SB;
            public string GMOFrm001BS;



        }
        public void Finally()
        {
           
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            DATAPROLOSAL PROLOSAL = new DATAPROLOSAL();
           
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            PROLOSAL.TaskId = applyTask.Task.TaskId;

            PROLOSAL.GMOFrm001AVG6 = applyTask.Task.CurrentDocument.Fields["GMOFrm001AVG6"].FieldValue.ToString().Trim();
            PROLOSAL.GMOFrm001SB = applyTask.Task.CurrentDocument.Fields["GMOFrm001SB"].FieldValue.ToString().Trim();
            


            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                if (!string.IsNullOrEmpty(PROLOSAL.TaskId))
                {
                    UPDATE_TB_WKF_TASK(PROLOSAL);
                }
            }

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public void UPDATE_TB_WKF_TASK(DATAPROLOSAL PROLOSAL)
        {
            decimal COUNTS = 0;
            string GMOFrm001BS = "";

            if (!string.IsNullOrEmpty(PROLOSAL.GMOFrm001SB)&& Convert.ToDecimal(PROLOSAL.GMOFrm001SB)>0)
            {
                COUNTS = Convert.ToDecimal(PROLOSAL.GMOFrm001SB);
            }
            else if(!string.IsNullOrEmpty(PROLOSAL.GMOFrm001AVG6) && Convert.ToDecimal(PROLOSAL.GMOFrm001AVG6) > 0)
            {
                COUNTS = Convert.ToDecimal(PROLOSAL.GMOFrm001AVG6);
            }

            if(COUNTS>=81)
            {
                GMOFrm001BS = "一級81分以上(6000元)";
            }
            else if(COUNTS<=80 && COUNTS>=76)
            {
                GMOFrm001BS = "二級76-80分(4000元)";
            }
            else if (COUNTS <= 75 && COUNTS >= 66)
            {
                GMOFrm001BS = "三級66-75分(3000元)";
            }
            else if (COUNTS <= 65 && COUNTS >= 56)
            {
                GMOFrm001BS = "四級56-65分(500元)";
            }
            else if (COUNTS <= 55 && COUNTS >= 45)
            {
                GMOFrm001BS = "五級45-55分(300元)";
            }
            else if (COUNTS <= 45 && COUNTS >= 35)
            {
                GMOFrm001BS = "六級35-45分(獎品乙份)";
            }


            try
            {   
                //宣告Task修改指定欄位內容

                TaskUtility taskUtility = new TaskUtility();

                /// <summary>
                /// 修改目前表單內容
                /// </summary>
                /// <param name="IsOptionalField">是否為外掛欄位</param>
                /// <param name="taskId">表單TASK_ID</param>
                /// <param name="fieldId">欄位代號</param>
                /// <param name="fieldValue">欄位內容(新的內容)</param>
                /// <param name="realValue">表單真實的值(for欄位式站點,傳入NULL代表不變動)</param>

                taskUtility.UpdateTaskContent(false, PROLOSAL.TaskId, "GMOFrm001BS", GMOFrm001BS, "");



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
