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

namespace TKUOF.TRIGGER.ASTI06
{
    //訂單的核準

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
                    UPDATE_ASTTC(TC001, TC002, FORMID, MODIFIER);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATE_ASTTC(string TC001, string TC002,string FORMID, string MODIFIER)
        {
            string TC015 = "Y";
            string TC032 = "N";
            string TC003 = DateTime.Now.ToString("yyyyMMdd");
            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();



            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"                                   
                                     UPDATE [TK].dbo.ASTTC
                                    SET TC003=@TC003
                                    ,TC015=@TC015
                                    ,FLAG=FLAG+1
                                    ,TC028=@TC028 
                                    ,TC032='N'
                                    ,UDF02=@UDF02
                                    ,MODIFIER=@MODIFIER
                                    ,MODI_DATE=MODI_DATE
                                    ,MODI_TIME=MODI_TIME 
                                    WHERE TC001=@TC001 AND TC002=@TC002

                                    UPDATE [TK].dbo.ASTMB
                                    set ASTMB.FLAG=ASTMB.FLAG+1
                                    ,MB021=TC007
                                    ,MB015=TC009+MB015
                                    ,MB022=TC010
                                    ,MB014=TC029+MB014
                                    ,MB049=TC030
                                    ,MB051=TC033 
                                    ,MB041=TC037
                                    ,MB074=TC007
                                    ,MB069=TC009
                                    ,MB068=TC029
                                    ,MB077=TC010
                                    ,MB063=TC038
                                    ,MB062=TC037
                                    ,MB064=TC039
                                    ,MB066=TC033
                                    ,MB065=TC030

                                    ,MODIFIER=@MODIFIER 
                                    ,MODI_DATE=@MODI_DATE 
                                    ,MODI_TIME=@MODI_TIME 
                                    FROM  [TK].dbo.ASTTC
                                    WHERE MB001=TC004
                                    AND TC001=@TC001 AND TC002=@TC002

                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TC001", SqlDbType.NVarChar).Value = TC001;
                    command.Parameters.Add("@TC002", SqlDbType.NVarChar).Value = TC002;
                    command.Parameters.Add("@TC003", SqlDbType.NVarChar).Value = TC003;
                    command.Parameters.Add("@TC015", SqlDbType.NVarChar).Value = TC015;
                    command.Parameters.Add("@TC028", SqlDbType.NVarChar).Value = MODIFIER;

                    command.Parameters.Add("@COMPANY", SqlDbType.NVarChar).Value = "TK";
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
