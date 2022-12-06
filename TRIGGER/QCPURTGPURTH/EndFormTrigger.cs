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

namespace TKUOF.TRIGGER.QCPURTGPURTH
{
    //品保檢驗進貨單明細

    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {
            
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            DataTable dt = new DataTable("QCHECK");
            dt.Columns.Add(new DataColumn("TH003", typeof(string)));
            dt.Columns.Add(new DataColumn("TH015", typeof(string)));
            dt.Columns.Add(new DataColumn("CHECK", typeof(string)));

            string TG001=null;
            string TG002 = null;
            string FORMID= null;
            string QCMAN = null;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TG001 = applyTask.Task.CurrentDocument.Fields["TG001"].FieldValue.ToString().Trim();
            TG002 = applyTask.Task.CurrentDocument.Fields["TG002"].FieldValue.ToString().Trim();
            FORMID = applyTask.FormNumber;
            //QCMAN= applyTask.Task.Applicant.UserName;

            //品保人員簽核就啟動，不用等整張表單簽完
            //核準
            //記錄TH003,TH015,CHECK
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve )
            {
                if (applyTask.SiteCode.Equals("QCCHECK"))
                {
                    if (!string.IsNullOrEmpty(TG001) && !string.IsNullOrEmpty(TG002))
                    {
                        foreach (XmlNode node in xmlDoc.SelectNodes("./Form/FormFieldValue/FieldItem[@fieldId='PURTH']/DataGrid/Row"))
                        {
                            string TH003 = node.SelectSingleNode("./Cell[@fieldId='TH003']").Attributes["fieldValue"].Value;
                            string TH015 = node.SelectSingleNode("./Cell[@fieldId='TH015']").Attributes["fieldValue"].Value;
                            string CHECK = node.SelectSingleNode("./Cell[@fieldId='CHECK']").Attributes["fieldValue"].Value;

                            //取得簽核人員
                            QCMAN = applyTask.Task.CurrentSigner.UserName;

                            DataRow dr = dt.NewRow();
                            dr["TH003"] = TH003;
                            dr["TH015"] = TH015;
                            dr["CHECK"] = CHECK;
                            dt.Rows.Add(dr);
                        }

                        if (dt.Rows.Count > 0)
                        {
                            UPDATEPURTGPURTH(TG001, TG002, FORMID, QCMAN, dt);

                            UPDATEPURTG(TG001, TG002);
                        }
                    }
                    

                }

            }
            //作廢
            else if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Cancel)
            {
              
            }
            //否決
            else if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Reject)
            {
               
            }
            //退簽
            else if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Return)
            {
              
            }

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }


        /// <summary>
        /// 更新進貨單的單身驗收、驗退數量、驗收狀況、驗收人員
        /// </summary>
        /// <param name="TG001"></param>
        /// <param name="TG002"></param>
        /// <param name="FORMID"></param>
        /// <param name="QCMAN"></param>
        /// <param name="dt"></param>
        public void UPDATEPURTGPURTH(string TG001, string TG002, string FORMID,string QCMAN,DataTable dt)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();

            if(dt.Rows.Count>0)
            {
                foreach(DataRow dr in dt.Rows)
                {

                    //設定TH028 檢驗狀態
                    //TH028='1' 待驗
                    //TH028='2' 合格
                    //TH028='3' 不良
                    string TH028 = "1";

                    if (dr["CHECK"].ToString().Equals("Y"))
                    {
                        TH028 = "2";
                    }
                    else
                    {
                        TH028 = "3";
                    }
                    queryString.AppendFormat(@"
                                            UPDATE [TK].dbo.PURTH
                                            SET TH015={1},UDF01='{2}',TH017=TH007-{4},TH028='{5}'
                                            WHERE TH001=@TG001 AND TH002=@TG002 AND TH003='{0}'
                                         ",  dr["TH003"].ToString(), Convert.ToDecimal(dr["TH015"].ToString()), dr["CHECK"].ToString()+','+ QCMAN+'-'+ FORMID, Convert.ToDecimal(dr["TH015"].ToString()), Convert.ToDecimal(dr["TH015"].ToString()), TH028);
                }
               
            }


            //queryString.AppendFormat(@"
                                      
            //                            ", FORMID);

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TG001", SqlDbType.NVarChar).Value = TG001;
                    command.Parameters.Add("@TG002", SqlDbType.NVarChar).Value = TG002;

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
        /// <summary>
        /// 更新PURTG的數量、原幣金額、原幣稅額、本幣金額、本幣稅額
        /// </summary>
        /// <param name="TG001"></param>
        /// <param name="TG002"></param>
        public void UPDATEPURTG(string TG001,string TG002)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();

            queryString.AppendFormat(@"

                                        UPDATE [TK].dbo.PURTH
                                        SET TH019=TEMP.TH019
                                        ,TH045=TEMP.TH045
                                        ,TH046=TEMP.TH046
                                        ,TH047=TEMP.TH047
                                        ,TH048=TEMP.TH048

                                        FROM 
                                        (
                                        SELECT TG001,TG002,TG010,TG030,TG008,TH003
                                        ,CASE WHEN TG010='1' THEN  ROUND(TH018*TH016,0)
                                        WHEN TG010='2' THEN  ROUND(TH018*TH016,0)
                                        WHEN TG010='3' THEN  ROUND(TH018*TH016,0)
                                        WHEN TG010='4' THEN  ROUND(TH018*TH016,0)
                                        WHEN TG010='5' THEN  ROUND(TH018*TH016,0)
                                        END  AS 'TH019'
                                        ,CASE WHEN TG010='1'THEN  ROUND(TH018*TH016/(1+TG030),0)
                                        WHEN TG010='2' THEN  ROUND(TH018*TH016,0)
                                        WHEN TG010='3' THEN  ROUND(TH018*TH016,0)
                                        WHEN TG010='4' THEN  ROUND(TH018*TH016,0)
                                        WHEN TG010='5' THEN  ROUND(TH018*TH016,0)
                                        END  AS 'TH045'
                                        ,CASE WHEN TG010='1' THEN  ROUND((TH018*TH016),0)-ROUND(TH018*TH016/(1+TG030),0)
                                        WHEN TG010='2' THEN ROUND(TH018*TH016*TG030,0)
                                        WHEN TG010='3' THEN 0 
                                        WHEN TG010='4' THEN 0
                                        WHEN TG010='5' THEN 0
                                        END  AS 'TH046'
                                        ,CASE WHEN TG010='1' THEN  ROUND(TH018*TH016*TG008/(1+TG030),0)
                                        WHEN TG010='2' THEN  ROUND(TH018*TH016*TG008,0)
                                        WHEN TG010='3' THEN  ROUND(TH018*TH016*TG008,0)
                                        WHEN TG010='4' THEN  ROUND(TH018*TH016*TG008,0)
                                        WHEN TG010='5' THEN  ROUND(TH018*TH016*TG008,0)
                                        END  AS 'TH047'
                                        ,CASE WHEN TG010='1' THEN  ROUND((TH018*TH016*TG008),0)-ROUND(TH018*TH016*TG008/(1+TG030),0)
                                        WHEN TG010='2' THEN ROUND(TH018*TH016*TG008*TG030,0)
                                        WHEN TG010='3' THEN 0 
                                        WHEN TG010='4' THEN 0
                                        WHEN TG010='5' THEN 0
                                        END  AS 'TH048'

                                        FROM [TK].dbo.PURTG,[TK].dbo.PURTH
                                        WHERE TG001=TH001 AND TG002=TH002
                                        AND TH001=@TG001 AND TH002=@TG002
                                        )
                                        AS TEMP
                                        WHERE PURTH.TH001=TEMP.TG001 AND PURTH.TH002=TEMP.TG002 AND PURTH.TH003=TEMP.TH003


                                        UPDATE [TK].dbo.PURTG
                                        SET TG026=TEMP.TG026
                                        ,TG017=TEMP.TG017
                                        ,TG019=TEMP.TG019
                                        ,TG028=TEMP.TG028
                                        ,TG031=TEMP.TG031
                                        ,TG032=TEMP.TG032
                                        FROM 
                                        (
                                        SELECT TG001,TG002
                                        ,(SELECT SUM(TH016) FROM [TK].dbo.PURTH WHERE TH001=TG001 AND TH002=TG002) TG026
                                        ,(SELECT SUM(TH045) FROM [TK].dbo.PURTH WHERE TH001=TG001 AND TH002=TG002) TG017
                                        ,(SELECT SUM(TH046) FROM [TK].dbo.PURTH WHERE TH001=TG001 AND TH002=TG002) TG019
                                        ,(SELECT SUM(TH045) FROM [TK].dbo.PURTH WHERE TH001=TG001 AND TH002=TG002) TG028
                                        ,(SELECT SUM(TH047) FROM [TK].dbo.PURTH WHERE TH001=TG001 AND TH002=TG002) TG031
                                        ,(SELECT SUM(TH048) FROM [TK].dbo.PURTH WHERE TH001=TG001 AND TH002=TG002) TG032
                                        FROM [TK].dbo.PURTG
                                        WHERE TG001=@TG001 AND TG002=@TG002
                                        ) AS TEMP
                                        WHERE PURTG.TG001=TEMP.TG001 AND PURTG.TG002=TEMP.TG002

                                        ");


            //queryString.AppendFormat(@"

            //                            ", FORMID);

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TG001", SqlDbType.NVarChar).Value = TG001;
                    command.Parameters.Add("@TG002", SqlDbType.NVarChar).Value = TG002;

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
