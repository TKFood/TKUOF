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

namespace TKUOF.TRIGGER.QCMOCTHMOCTI
{
    //品保檢驗託外進貨單明細

    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {
            
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            DataTable dt = new DataTable("QCHECK");
            dt.Columns.Add(new DataColumn("TI003", typeof(string)));
            dt.Columns.Add(new DataColumn("TI019", typeof(string)));
            dt.Columns.Add(new DataColumn("CHECK", typeof(string)));

            string TH001= null;
            string TH002 = null;
            string FORMID= null;
            string QCMAN = null;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TH001 = applyTask.Task.CurrentDocument.Fields["TH001"].FieldValue.ToString().Trim();
            TH002 = applyTask.Task.CurrentDocument.Fields["TH002"].FieldValue.ToString().Trim();
            FORMID = applyTask.FormNumber;
            //QCMAN= applyTask.Task.Applicant.UserName;

            //品保人員簽核就啟動，不用等整張表單簽完
            //核準
            //記錄TH003,TH015,CHECK
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve )
            {
                if (applyTask.SiteCode.Equals("QCCHECK"))
                {
                    if (!string.IsNullOrEmpty(TH001) && !string.IsNullOrEmpty(TH002))
                    {
                        foreach (XmlNode node in xmlDoc.SelectNodes("./Form/FormFieldValue/FieldItem[@fieldId='MOCTI']/DataGrid/Row"))
                        {
                            string TI003 = node.SelectSingleNode("./Cell[@fieldId='TI003']").Attributes["fieldValue"].Value;
                            string TI019 = node.SelectSingleNode("./Cell[@fieldId='TI019']").Attributes["fieldValue"].Value;
                            string CHECK = node.SelectSingleNode("./Cell[@fieldId='CHECK']").Attributes["fieldValue"].Value;

                            //取得簽核人員
                            QCMAN = applyTask.Task.CurrentSigner.UserName;

                            DataRow dr = dt.NewRow();
                            dr["TI003"] = TI003;
                            dr["TI019"] = TI019;
                            dr["CHECK"] = CHECK;
                            dt.Rows.Add(dr);
                        }

                        if (dt.Rows.Count > 0)
                        {
                            UPDATE_MOCTG_MOCTI(TH001, TH002, FORMID, QCMAN, dt);

                            UPDATE_MOCTG(TH001, TH002);
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
        public void UPDATE_MOCTG_MOCTI(string TH001, string TH002, string FORMID,string QCMAN,DataTable dt)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();

            if(dt.Rows.Count>0)
            {
                foreach(DataRow dr in dt.Rows)
                {

                    //設定TI035 檢驗狀態
                    //TI035='1' 待驗
                    //TI035='2' 合格
                    //TI035='3' 不良
                    string TI035 = "1";

                    if (dr["CHECK"].ToString().Equals("Y"))
                    {
                        TI035 = "2";
                    }
                    else
                    {
                        TI035 = "3";
                    }
                    queryString.AppendFormat(@"
                                            UPDATE [TK].dbo.MOCTI
                                            SET TI019={1},TI020={1},UDF01='{2}',TI022=TI007-{4},TI035='{5}'
                                            WHERE TI001=@TI001 AND TI002=@TI002 AND TI003='{0}'
                                         ",  dr["TI003"].ToString(), Convert.ToDecimal(dr["TI019"].ToString()), dr["CHECK"].ToString()+','+ QCMAN+'-'+ FORMID, Convert.ToDecimal(dr["TI019"].ToString()), Convert.ToDecimal(dr["TI019"].ToString()), TI035);
                }
               
            }


            //queryString.AppendFormat(@"
                                      
            //                            ", FORMID);

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TI001", SqlDbType.NVarChar).Value = TH001;
                    command.Parameters.Add("@TI002", SqlDbType.NVarChar).Value = TH002;

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
        public void UPDATE_MOCTG(string TH001,string TH002)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();

            queryString.AppendFormat(@"
                                        UPDATE [TK].dbo.MOCTI
                                        SET 
                                        TI025=TEMP.TI025
                                        ,TI044=TEMP.TI044
                                        ,TI045=TEMP.TI045
                                        ,TI046=TEMP.TI046
                                        ,TI047=TEMP.TI047

                                        FROM 
                                        (
                                        SELECT TH001,TH002,TH015,TH030,TH008,TI003
                                        ,CASE WHEN TH015='1' THEN  ROUND(TI024*TI020,0)
                                        WHEN TH015='2' THEN  ROUND(TI024*TI020,0)
                                        WHEN TH015='3' THEN  ROUND(TI024*TI020,0)
                                        WHEN TH015='4' THEN  ROUND(TI024*TI020,0)
                                        WHEN TH015='5' THEN  ROUND(TI024*TI020,0)
                                        END  AS 'TI025'
                                        ,CASE WHEN TH015='1'THEN  ROUND(TI024*TI020/(1+TH030),0)
                                        WHEN TH015='2' THEN  ROUND(TI024*TI020,0)
                                        WHEN TH015='3' THEN  ROUND(TI024*TI020,0)
                                        WHEN TH015='4' THEN  ROUND(TI024*TI020,0)
                                        WHEN TH015='5' THEN  ROUND(TI024*TI020,0)
                                        END  AS 'TI044'
                                        ,CASE WHEN TH015='1' THEN  ROUND((TI024*TI020),0)-ROUND(TI024*TI020/(1+TH030),0)
                                        WHEN TH015='2' THEN ROUND(TI024*TI020*TH030,0)
                                        WHEN TH015='3' THEN 0 
                                        WHEN TH015='4' THEN 0
                                        WHEN TH015='5' THEN 0
                                        END  AS 'TI045'
                                        ,CASE WHEN TH015='1' THEN  ROUND(TI024*TI020*TH008/(1+TH030),0)
                                        WHEN TH015='2' THEN  ROUND(TI024*TI020*TH008,0)
                                        WHEN TH015='3' THEN  ROUND(TI024*TI020*TH008,0)
                                        WHEN TH015='4' THEN  ROUND(TI024*TI020*TH008,0)
                                        WHEN TH015='5' THEN  ROUND(TI024*TI020*TH008,0)
                                        END  AS 'TI046'
                                        ,CASE WHEN TH015='1' THEN  ROUND((TI024*TI020*TH008),0)-ROUND(TI024*TI020*TH008/(1+TH030),0)
                                        WHEN TH015='2' THEN ROUND(TI024*TI020*TH008*TH030,0)
                                        WHEN TH015='3' THEN 0 
                                        WHEN TH015='4' THEN 0
                                        WHEN TH015='5' THEN 0
                                        END  AS 'TI047'

                                        FROM [TK].dbo.MOCTH,[TK].dbo.MOCTI
                                        WHERE TH001=TI001 AND TH002=TI002
                                        AND TI001=@TH001 AND TI002=@TH002
                                        )
                                        AS TEMP
                                        WHERE MOCTI.TI001=TEMP.TH001 AND MOCTI.TI002=TEMP.TH002 AND MOCTI.TI003=TEMP.TI003


                                        UPDATE [TK].dbo.MOCTH
                                        SET TH022=TEMP.TH022
                                        ,TH018=TEMP.TH018
                                        ,TH019=TEMP.TH019
                                        ,TH027=TEMP.TH027
                                        ,TH031=TEMP.TH031
                                        ,TH032=TEMP.TH032
                                        ,TH020=TEMP.TH020

                                        FROM 
                                        (
                                        SELECT TH001,TH002
                                        ,(SELECT SUM(TI020) FROM [TK].dbo.MOCTI WHERE TH001=TI001 AND TH002=TI002) TH022
                                        ,(SELECT SUM(TI025) FROM [TK].dbo.MOCTI WHERE TH001=TI001 AND TH002=TI002) TH018
                                        ,(SELECT SUM(TI026) FROM [TK].dbo.MOCTI WHERE TH001=TI001 AND TH002=TI002) TH019
                                        ,(SELECT SUM(TI025) FROM [TK].dbo.MOCTI WHERE TH001=TI001 AND TH002=TI002) TH027
                                        ,(SELECT SUM(TI046) FROM [TK].dbo.MOCTI WHERE TH001=TI001 AND TH002=TI002) TH031
                                        ,(SELECT SUM(TI047) FROM [TK].dbo.MOCTI WHERE TH001=TI001 AND TH002=TI002) TH032
                                        ,(SELECT SUM(TI045) FROM [TK].dbo.MOCTI WHERE TH001=TI001 AND TH002=TI002) TH020
                                        FROM [TK].dbo.MOCTH
                                        WHERE TH001=@TH001 AND TH002=@TH002
                                        ) AS TEMP
                                        WHERE MOCTH.TH001=TEMP.TH001 AND MOCTH.TH002=TEMP.TH002

                                        ");


            //queryString.AppendFormat(@"

            //                            ", FORMID);

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TH001", SqlDbType.NVarChar).Value = TH001;
                    command.Parameters.Add("@TH002", SqlDbType.NVarChar).Value = TH002;

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
