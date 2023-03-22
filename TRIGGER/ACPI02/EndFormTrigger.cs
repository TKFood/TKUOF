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

namespace TKUOF.TRIGGER.ACPI02
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
                    UPDATE_ACPI02(TA001, TA002, FORMID, MODIFIER);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATE_ACPI02(string TA001, string TA002, string FORMID, string MODIFIER)
        {
            string TA024 = "Y";
            string TA035 = MODIFIER;
            string TA044 = "N";
            string TB012 = "Y";
            string TB001 = TA001;
            string TB002 = TA002;
            string TH031 = "Y";
            string TJ021 = "Y";
            string TI038 = "Y";
            string TL025 = "Y";
            string TA026 = "Y";
            string TA051= DateTime.Now.ToString("yyyyMMdd");
            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();



            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"       
                                    
                                        UPDATE [test0923].dbo.ACPTA
                                        SET 
                                        TA024=@TA024 
                                        ,TA035=@TA035 
                                        ,TA044=@TA044 
                                        ,UDF02=@UDF02
                                        ,FLAG=FLAG+1
                                        ,MODIFIER=@MODIFIER 
                                        ,MODI_DATE=@MODI_DATE
                                        ,MODI_TIME=@MODI_TIME
                                        WHERE TA001=@TA001 AND TA002=@TA002

                                        UPDATE [TK].dbo.ACPTB
                                        SET
                                        TB012=@TB012
                                        ,FLAG=FLAG+1
                                        ,MODIFIER=@MODIFIER 
                                        ,MODI_DATE=@MODI_DATE
                                        ,MODI_TIME=@MODI_TIME
                                        WHERE TB001=@TB001 AND TB002=@TB002


                                        UPDATE [test0923].dbo.PURTH
                                        SET 
                                        TH031=@TH031 
                                        ,TH039=TB001
                                        ,TH040=TB002
                                        ,TH041=TB003
                                        ,PURTH.FLAG=PURTH.FLAG+1 
                                        ,MODIFIER=@MODIFIER
                                        ,MODI_DATE=@MODI_DATE 
                                        ,MODI_TIME=@MODI_TIME
                                        FROM [test0923].dbo.ACPTB
                                        WHERE 1=1
                                        AND TB004='1'
                                        AND TH001=TB005 AND TH002=TB006 AND TH003=TB007
                                        AND TB001=@TB001 AND TB002=@TB002



                                        UPDATE [test0923].dbo.PURTJ
                                        SET
                                        TJ021=@TJ021
                                        ,TJ025=TB001
                                        ,TJ026=TB002
                                        ,TJ027=TB003
                                        ,PURTJ.FLAG=PURTJ.FLAG+1
                                        ,MODIFIER=@MODIFIER
                                        ,MODI_DATE=@MODI_DATE
                                        ,MODI_TIME=@MODI_TIME 
                                        FROM [test0923].dbo.ACPTB
                                        WHERE 1=1
                                        AND TB004='2'
                                        AND TJ001=TB005 AND TJ002=TB006 AND TJ003=TB007
                                        AND TB001=@TB001 AND TB002=@TB002

                                        UPDATE [test0923].dbo.MOCTI
                                        SET
                                        TI038=@TI038
                                        ,TI029=TB001
                                        ,TI030=TB002
                                        ,TI031=TB003
                                        ,MOCTI.FLAG=MOCTI.FLAG+1
                                        ,MODIFIER=@MODIFIER
                                        ,MODI_DATE=@MODI_DATE
                                        ,MODI_TIME=@MODI_TIME
                                        FROM [test0923].dbo.ACPTB
                                        WHERE 1=1
                                        AND TB004='3'
                                        AND TI001=TB005 AND TI002=TB006 AND TI003=TB007
                                        AND TB001=@TB001 AND TB002=@TB002


                                        UPDATE [test0923].dbo.MOCTL
                                        SET
                                        TL025=@TL025
                                        ,TL019=TB001
                                        ,TL020=TB002
                                        ,TL021=TB003
                                        ,MOCTL.FLAG=MOCTL.FLAG+1
                                        ,MODIFIER=@MODIFIER
                                        ,MODI_DATE=@MODI_DATE 
                                        ,MODI_TIME=@MODI_TIME
                                        FROM [test0923].dbo.ACPTB
                                        WHERE 1=1
                                        AND TB004='4'
                                        AND TL001=TB005 AND TL002=TB006 AND TL003=TB007
                                        AND TB001=@TB001 AND TB002=@TB002


                                        UPDATE [test0923].dbo.ACPTA
                                        SET 
                                        TA030=TB017+TB018
                                        ,TA026=@TA026
                                        ,TA043=TB018
                                        ,TA048=TB017+TB018
                                        ,TA051=@TA051
                                        FROM [test0923].dbo.ACPTB
                                        WHERE 1=1
                                        AND TB004='A'
                                        AND TA001=TB005 AND TA002=TB006
                                        AND TB001=@TB001 AND TB002=@TB002




                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TA001", SqlDbType.NVarChar).Value = TA001;
                    command.Parameters.Add("@TA002", SqlDbType.NVarChar).Value = TA002;

                    command.Parameters.Add("@TA024", SqlDbType.NVarChar).Value = TA024;
                    command.Parameters.Add("@TA035", SqlDbType.NVarChar).Value = TA035;
                    command.Parameters.Add("@TA044", SqlDbType.NVarChar).Value = TA044;
                    command.Parameters.Add("@TB012", SqlDbType.NVarChar).Value = TB012;
                    command.Parameters.Add("@TB001", SqlDbType.NVarChar).Value = TB001;
                    command.Parameters.Add("@TB002", SqlDbType.NVarChar).Value = TB002;
                    command.Parameters.Add("@TH031", SqlDbType.NVarChar).Value = TH031;
                    command.Parameters.Add("@TJ021", SqlDbType.NVarChar).Value = TJ021;
                    command.Parameters.Add("@TI038", SqlDbType.NVarChar).Value = TI038;
                    command.Parameters.Add("@TL025", SqlDbType.NVarChar).Value = TL025;
                    command.Parameters.Add("@TA026", SqlDbType.NVarChar).Value = TA026;
                    command.Parameters.Add("@TA051", SqlDbType.NVarChar).Value = TA051;
                   

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
