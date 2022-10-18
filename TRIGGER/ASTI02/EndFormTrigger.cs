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

namespace TKUOF.TRIGGER.ASTI02
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
            string MB042 = "AC01";
            string MB043 = null;
            string FORMID = null;
            string MODIFIER = null;
            UserUCO userUCO = new UserUCO();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            MB001 = applyTask.Task.CurrentDocument.Fields["MB001"].FieldValue.ToString().Trim();     
          
            FORMID = applyTask.FormNumber;
            //MODIFIER = applyTask.Task.Applicant.Account;
      
            //取得簽核人工號
            EBUser ebUser = userUCO.GetEBUser(applyTask.Task.CurrentSite.CurrentNode.ActualSignerId);     
            MODIFIER = ebUser.Account;

            //取得MB042的單號
            MB043 = GETMAXNO(MB042,DateTime.Now.ToString("yyyyMMdd"));


            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(MB001) && !string.IsNullOrEmpty(MB042) && !string.IsNullOrEmpty(MB043))
                {
                    UPDATEASTMBASTMC(MB001, FORMID, MODIFIER, MB042, MB043);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATEASTMBASTMC(string MB001, string FORMID, string MODIFIER, string MB042, string MB043)
        {
            string TC027 = "Y";
            string TC048 = "N";
            string TD021 = "Y";
            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();



            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"                                   
                                        UPDATE [TK].dbo.ASTMB
                                        SET MB047=@MB047,MB048=@MB048
                                        ,MB042=@MB042,MB043=@MB043
                                        ,MB039=@MB039,MB050=@MB050
                                        ,FLAG=FLAG+1,COMPANY=@COMPANY,MODIFIER=@MODIFIER,MODI_DATE=@MODI_DATE,MODI_TIME=@MODI_TIME
                                        WHERE MB001=@MB001




                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@MB001", SqlDbType.NVarChar).Value = MB001;
                    command.Parameters.Add("@MB047", SqlDbType.NVarChar).Value = DateTime.Now.ToString("yyyyMMdd");
                    command.Parameters.Add("@MB048", SqlDbType.NVarChar).Value = MODIFIER;
                    command.Parameters.Add("@MB042", SqlDbType.NVarChar).Value = MB042;
                    command.Parameters.Add("@MB043", SqlDbType.NVarChar).Value = MB043;
                    command.Parameters.Add("@MB039", SqlDbType.NVarChar).Value = "Y";
                    command.Parameters.Add("@MB050", SqlDbType.NVarChar).Value = "N";
                    command.Parameters.Add("@COMPANY", SqlDbType.NVarChar).Value = "TK";
                    command.Parameters.Add("@MODIFIER", SqlDbType.NVarChar).Value = MODIFIER;
                    command.Parameters.Add("@MODI_DATE", SqlDbType.NVarChar).Value = DateTime.Now.ToString("yyyyMMdd");
                    command.Parameters.Add("@MODI_TIME", SqlDbType.NVarChar).Value = DateTime.Now.ToString("HH:mm:ss");



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


        public string GETMAXNO(string TC001,string TC003)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["CHIYUconnectionstring"].ToString();
            Ede.Uof.Utility.Data.DatabaseHelper m_db = new Ede.Uof.Utility.Data.DatabaseHelper(connectionString);

            string cmdTxt = @" 
                            SELECT ISNULL(MAX(TC002),'00000000000') AS TA002
                            FROM  [TK].dbo.ASTTC
                            WHERE  TC001=@TC001 AND TC003=@TC003
                             ";

            m_db.AddParameter("@TC001", TC001);
            m_db.AddParameter("@TC003", TC003);


            DataTable dt = new DataTable();

            dt.Load(m_db.ExecuteReader(cmdTxt));

            if (dt.Rows.Count > 0)
            {
                string TA002 = SETTA002(dt.Rows[0]["TA002"].ToString());
                return TA002;
            }
            else
            {
                return null;
            }
        }

        public string SETTA002(string TA002)
        {
            DateTime dt = DateTime.Now;

            if (TA002.Equals("00000000000"))
            {
                return dt.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(TA002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return dt.ToString("yyyyMMdd") + temp.ToString();
            }

        }
    }
}
