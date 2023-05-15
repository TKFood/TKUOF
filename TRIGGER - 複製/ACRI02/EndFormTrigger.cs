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
using Ede.Uof.EIP.SystemInfo;

namespace TKUOF.TRIGGER.ACRI02
{
    //ACRI02.結帳單 的核準


    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {

        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TA001 = null;
            string TA002 = null;
            string FORMID = null;
            string MODIFIER = null;
            UserUCO userUCO = new UserUCO();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TA001 = applyTask.Task.CurrentDocument.Fields["TA001"].FieldValue.ToString().Trim();
            TA002 = applyTask.Task.CurrentDocument.Fields["TA002"].FieldValue.ToString().Trim();

            FORMID = applyTask.FormNumber;
            //MODIFIER = applyTask.Task.Applicant.Account;

            //取得簽核人工號
            EBUser ebUser = userUCO.GetEBUser(Current.UserGUID);
            MODIFIER = ebUser.Account;

           
            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(TA001) && !string.IsNullOrEmpty(TA002) )
                {
                    UPDATE_ACRI02(TA001, TA002, FORMID, MODIFIER);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATE_ACRI02(string TA001, string TA002, string FORMID, string MODIFIER)
        {
            string TA025 = "Y";
            string TA048 = "N";
            string TA039 = MODIFIER;
            string TB012 = "Y";
            string TH026 = "Y";
            string TJ024 = "Y";
            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();



            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"       
                                    UPDATE [test0923].dbo.ACRTA
                                    SET 
                                    TA025=@TA025 
                                    ,TA048=@TA048 
                                    ,TA039=@MODIFIER 
                                    ,FLAG=FLAG+1
                                    ,MODIFIER=@MODIFIER 
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME 
                                    WHERE TA001=@TA001 AND TA002=@TA002

                                    UPDATE [test0923].dbo.ACRTB
                                    SET
                                     TB012=@TB012 
                                    ,FLAG=FLAG+1
                                    ,MODIFIER=@MODIFIER 
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME 
                                    WHERE TB001=@TB001 AND TB002=@TB002


                                    UPDATE [test0923].dbo.COPTH
                                    SET
                                    TH026=@TH026 
                                    ,TH027=TB001
                                    ,TH028=TB002
                                    ,TH029=TB003
                                    FROM [test0923].dbo.ACRTB
                                    WHERE TB001=@TB001 AND TB002=@TB002
                                    AND TB005=TH001 AND TB006=TH002 

                                    UPDATE [test0923].dbo.COPTH
                                    SET
                                    TH026=@TH026 
                                    ,TH027=TB001
                                    ,TH028=TB002
                                    ,TH029=TB003
                                    FROM [test0923].dbo.ACRTB
                                    WHERE TB001=@TB001 AND TB002=@TB002
                                    AND TB005=TH001 AND TB006=TH002 AND TB007=TH003

                                    UPDATE [test0923].dbo.COPTJ
                                    SET
                                    TJ024=TJ024 
                                    ,TJ025=TB001
                                    ,TJ026=TB002
                                    ,TJ027=TB003
                                    FROM [test0923].dbo.ACRTB
                                    WHERE TB001=@TB001 AND TB002=@TB002
                                    AND TB005=TJ001 AND TB006=TJ002 

                                    UPDATE [test0923].dbo.COPTJ
                                    SET
                                    TJ024=TJ024
                                    ,TJ025=TB001
                                    ,TJ026=TB002
                                    ,TJ027=TB003
                                    FROM [test0923].dbo.ACRTB
                                    WHERE TB001=@TB001 AND TB002=@TB002
                                    AND TB005=TJ001 AND TB006=TJ002 AND TB007=TJ003
                                  
                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TA001", SqlDbType.NVarChar).Value = TA001;
                    command.Parameters.Add("@TA002", SqlDbType.NVarChar).Value = TA002;

                    command.Parameters.Add("@TB001", SqlDbType.NVarChar).Value = TA001;
                    command.Parameters.Add("@TB002", SqlDbType.NVarChar).Value = TA002;
                    command.Parameters.Add("@TA025", SqlDbType.NVarChar).Value = TA025;
                    command.Parameters.Add("@TA048", SqlDbType.NVarChar).Value = TA048;
                    command.Parameters.Add("@TA039", SqlDbType.NVarChar).Value = TA039;
                    command.Parameters.Add("@TB012", SqlDbType.NVarChar).Value = TB012;
                    command.Parameters.Add("@TH026", SqlDbType.NVarChar).Value = TH026;                    
                    command.Parameters.Add("@TJ024", SqlDbType.NVarChar).Value = TJ024;

                    command.Parameters.Add("@MODIFIER", SqlDbType.NVarChar).Value = MODIFIER;
                    command.Parameters.Add("@MODI_DATE", SqlDbType.NVarChar).Value = DateTime.Now.ToString("yyyyMMdd");
                    command.Parameters.Add("@MODI_TIME", SqlDbType.NVarChar).Value = DateTime.Now.ToString("HH:mm:ss");
                    command.Parameters.Add("@UDF02", SqlDbType.NVarChar).Value =FORMID;



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
