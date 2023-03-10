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

namespace TKUOF.TRIGGER.ASTI10
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
            string TC028 = MODIFIER;
            string TD009 = "Y";
            //string TC003 = DateTime.Now.ToString("yyyyMMdd");
            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();



            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"                                   
                                     UPDATE [TK].dbo.ASTTC
                                    SET 
                                    TC015=@TC015
                                    ,TC028=@TC028 
                                    ,TC032=@TC032 
                                    ,FLAG=FLAG+2
                                    ,MODIFIER=@MODIFIER 
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME
                                    WHERE TC001=@TC001 AND TC002=@TC002

                                    UPDATE [TK].dbo.ASTTD
                                    SET 
                                    TD009=@TD009
                                    ,FLAG=FLAG+1
                                    ,MODIFIER=@MODIFIER 
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME 
                                    WHERE TD001=@TD001 AND TD002=@TD002

                                    UPDATE [TK].dbo.ASTMC
                                    SET MC004=TD005
                                    FROM [TK].dbo.ASTTC,[TK].dbo.ASTTD
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TC004=MC001 AND TD003=MC002 AND TD004=MC002
                                    AND TC001=@TC001 AND TC002=@TC002


                                   UPDATE [TK].dbo.ASTMB
                                    SET
                                   ASTMB.FLAG=ASTMB.FLAG+1 
                                    ,MB012=TC005 
                                    ,MB015=TC009 
                                    ,MB020=TC006 
                                    ,MB021=TC007 
                                    ,MB022=TC010 
                                    ,MB029=TC008 
                                    ,MB058=TC036 
                                    ,MB049=TC030 
                                    ,MB051=TC033  
                                    ,MB025=TC059
                                    ,MB014=TC035
                                    ,MB041=TC038
                                    ,MB027=TC039
                                    ,MB026=TC037
                                    ,MB081=TC038
                                    ,MB073=TC075
                                    ,MB074=TC076
                                    ,MB075=TC077
                                    ,MB076=TC068
                                    ,MB066=TC079
                                    ,MB069=TC073
                                    ,MB068=TC074
                                    ,MB077=TC078
                                    ,MB063=TC038
                                    ,MB062=TC037
                                    ,MB064=TC039
                                    ,MB065=TC030
                                    ,MB061=TC059
                                    ,MB078=TC092
                                    ,MB079=TC093
                                    ,MB023=TC097
                                    ,MB059=TC099
                                    ,MB082=TC061 
                                    ,MB019=TC007 
                                    ,MB072=TC075
                                    ,MODIFIER=@MODIFIER
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME 
                                    FROM [TK].dbo.ASTTC
                                    WHERE TC004=MB001
                                    AND TC001=@TC001 AND TC002=@TC002

                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TC001", SqlDbType.NVarChar).Value = TC001;
                    command.Parameters.Add("@TC002", SqlDbType.NVarChar).Value = TC002;
         ;
                    command.Parameters.Add("@TC015", SqlDbType.NVarChar).Value = TC015;
                    command.Parameters.Add("@TC028", SqlDbType.NVarChar).Value = TC028;
                    command.Parameters.Add("@TC032", SqlDbType.NVarChar).Value = TC032;

                    command.Parameters.Add("@TD001", SqlDbType.NVarChar).Value = TC001;
                    command.Parameters.Add("@TD002", SqlDbType.NVarChar).Value = TC002;

                    command.Parameters.Add("@TD009", SqlDbType.NVarChar).Value = TD009;

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
