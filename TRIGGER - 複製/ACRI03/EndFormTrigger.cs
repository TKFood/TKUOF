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

namespace TKUOF.TRIGGER.ACRI03
{
    //ACRI02.結帳單 的核準


    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {

        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TC001 = null;
            string TC002 = null;
            string FORMID = null;
            string MODIFIER = null;
            UserUCO userUCO = new UserUCO();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TC001 = applyTask.Task.CurrentDocument.Fields["TC001"].FieldValue.ToString().Trim();
            TC002 = applyTask.Task.CurrentDocument.Fields["TC002"].FieldValue.ToString().Trim();

            FORMID = applyTask.FormNumber;
            //MODIFIER = applyTask.Task.Applicant.Account;

            //取得簽核人工號
            EBUser ebUser = userUCO.GetEBUser(Current.UserGUID);
            MODIFIER = ebUser.Account;

           
            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002) )
                {
                    UPDATE_ACRI03(TC001, TC002, FORMID, MODIFIER);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATE_ACRI03(string TC001, string TC002, string FORMID, string MODIFIER)
        {
            string TC008 = "Y";
            string TD020 = "Y";
            string TA027 = "Y";
            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();



            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"       
                                    
                                    UPDATE [test0923].dbo.ACRTC
                                    SET 
                                    TC008=@TC008
                                    ,TC018=@MODIFIER 
                                    ,MODIFIER=@MODIFIER 
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME
                                    WHERE TC001=@TC001 AND TC002=@TC002 

                                    UPDATE [test0923].dbo.ACRTD
                                    SET
                                    TD020=@TD020
                                    ,MODIFIER=@MODIFIER 
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME
                                    WHERE TD001=@TD001 AND TD002=@TD002

                                    UPDATE [test0923].dbo.ACRTA
                                    SET 
                                    TA027=@TA027
                                    ,TA031=TA031+TD014
                                    ,TA058=TA058+TD015
                                    ,TA062=@MODI_DATE
                                    FROM [test0923].dbo.ACRTD
                                    WHERE TD001=@TD001 AND TD002=TD002
                                    AND TD013=TD014
                                    AND TD006=TA001 AND TD007=TA002

                                    UPDATE [test0923].dbo.ACRTA
                                    SET 
                                    TA027='N'
                                    ,TA031=TA031+TD014
                                    ,TA058=TA058+TD015
                                    ,TA062=''
                                    FROM [test0923].dbo.ACRTD
                                    WHERE TD001=@TD001 AND TD002=TD002
                                    AND TD013<>TD014
                                    AND TD006=TA001 AND TD007=TA002

                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TC001", SqlDbType.NVarChar).Value = TC001;
                    command.Parameters.Add("@TC002", SqlDbType.NVarChar).Value = TC002;

                    command.Parameters.Add("@TD001", SqlDbType.NVarChar).Value = TC001;
                    command.Parameters.Add("@TD002", SqlDbType.NVarChar).Value = TC002;

                    command.Parameters.Add("@TC008", SqlDbType.NVarChar).Value = TC008;
                    command.Parameters.Add("@TD020", SqlDbType.NVarChar).Value = TD020;
                    command.Parameters.Add("@TA027", SqlDbType.NVarChar).Value = TA027;

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
