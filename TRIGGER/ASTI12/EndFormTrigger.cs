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

namespace TKUOF.TRIGGER.ASTI12
{
    //訂單的核準
 

    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {

        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TE001 = null;
            string TE002 = null;
            string FORMID = null;
            string MODIFIER = null;
            UserUCO userUCO = new UserUCO();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TE001 = applyTask.Task.CurrentDocument.Fields["TE001"].FieldValue.ToString().Trim();
            TE002 = applyTask.Task.CurrentDocument.Fields["TE002"].FieldValue.ToString().Trim();

            FORMID = applyTask.FormNumber;
            //MODIFIER = applyTask.Task.Applicant.Account;

            //取得簽核人工號
            EBUser ebUser = userUCO.GetEBUser(Current.UserGUID);
            MODIFIER = ebUser.Account;

           
            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(TE001) && !string.IsNullOrEmpty(TE002) )
                {
                    UPDATE_ASTMC(TE001, TE002, FORMID, MODIFIER);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATE_ASTMC(string TE001, string TE002, string FORMID, string MODIFIER)
        {
            string TE006 = "Y";
            string TE010 = "N";
            string TE009 = MODIFIER;
            string TF009 = "Y";
            //string TC003 = DateTime.Now.ToString("yyyyMMdd");
            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();



            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"                                   
                                    UPDATE [TK].dbo.ASTTE
                                    SET
                                    TE006=@TE006 
                                    ,TE009=@TE009
                                    ,TE010=@TE010
                                    ,FLAG=FLAG+1
                                    ,MODIFIER=@MODIFIER
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME
                                    ,UDF02=@UDF02
                                    WHERE TE001=@TE001 AND TE002=@TE002

                                    
  
                                    UPDATE [TK].dbo.ASTTF
                                    SET
                                    TF009=@TF009 
                                    ,FLAG=FLAG+1
                                    ,MODIFIER=@MODIFIER
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME 
                                    WHERE TF001=@TF001 AND TF002=@TF002

                                    DELETE  [TK].dbo.ASTMC
                                    WHERE Replace(MC001+MC002+MC003,' ','') IN (
                                    SELECT Replace(TF003+TF004+TF005,' ','')
                                    FROM [TK].dbo.ASTTF
                                    WHERE TF001=@TF001 AND TF002=@TF002
                                    )


                                    INSERT INTO [TK].dbo.ASTMC
                                    (MC001 ,MC002 ,MC003 ,MC004 ,MC005 ,MC006, 
                                    COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
                                    CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME)
                                    SELECT TF003,TF104,TF105,TF006,TF008,TF107
                                    ,N'TK',TF105,MF007,CONVERT(nvarchar,GETDATE(),112),0,CONVERT(nvarchar,GETDATE(),108),N'',N'P004',N'ASTI12'
                                    FROM [TK].dbo.ASTTF
                                    LEFT JOIN [TK].dbo.ADMMF ON MF001=TF105
                                    WHERE TF001=@TF001 AND TF002=@TF002



                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TE001", SqlDbType.NVarChar).Value = TE001;
                    command.Parameters.Add("@TE002", SqlDbType.NVarChar).Value = TE002;

                    command.Parameters.Add("@TE006", SqlDbType.NVarChar).Value = TE006;
                    command.Parameters.Add("@TE009", SqlDbType.NVarChar).Value = TE009;
                    command.Parameters.Add("@TE010", SqlDbType.NVarChar).Value = TE010;
                    command.Parameters.Add("@TF009", SqlDbType.NVarChar).Value = TF009;

                    command.Parameters.Add("@TF001", SqlDbType.NVarChar).Value = TE001;
                    command.Parameters.Add("@TF002", SqlDbType.NVarChar).Value = TE002;


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
