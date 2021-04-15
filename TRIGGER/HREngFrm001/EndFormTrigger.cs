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

namespace TKUOF.TRIGGER.HREngFrm001
{
    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public class DATAHREngFrm001
        {
            public string TaskId;

            public string HREngFrm001PIR;
            public string HREngFrm001OU;
            public string HREngFrm001CN;

            public string HREngFrm001SN;
            public string HREngFrm001Date;
            public string HREngFrm001User;
            public string HREngFrm001UsrDpt;
            public string HREngFrm001Rank;
            public string HREngFrm001OutDate;
            public string HREngFrm001Location;
            public string HREngFrm001Agent;
            public string HREngFrm001Transp;
            public string HREngFrm001LicPlate;
            public string HREngFrm001Cause;
            public string HREngFrm001DefOutTime;
            public string HREngFrm001FF;
            public string HREngFrm001OutTime;
            public string HREngFrm001DefBakTime;
            public string HREngFrm001CH;
            public string HREngFrm001BakTime;
            public string CRADNO;


        }
        public void Finally()
        {
           
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            DATAHREngFrm001 HREngFrm001 = new DATAHREngFrm001();

           
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            HREngFrm001.TaskId = applyTask.Task.TaskId;

            HREngFrm001.HREngFrm001PIR = applyTask.Task.CurrentDocument.Fields["HREngFrm001PIR"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001OU = applyTask.Task.CurrentDocument.Fields["HREngFrm001OU"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001CN = applyTask.Task.CurrentDocument.Fields["HREngFrm001CN"].FieldValue.ToString().Trim();

            HREngFrm001.HREngFrm001SN = applyTask.Task.CurrentDocument.Fields["HREngFrm001SN"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001Date = applyTask.Task.CurrentDocument.Fields["HREngFrm001Date"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001User = applyTask.Task.CurrentDocument.Fields["HREngFrm001User"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001UsrDpt = applyTask.Task.CurrentDocument.Fields["HREngFrm001UsrDpt"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001Rank = applyTask.Task.CurrentDocument.Fields["HREngFrm001Rank"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001OutDate = applyTask.Task.CurrentDocument.Fields["HREngFrm001OutDate"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001Location = applyTask.Task.CurrentDocument.Fields["HREngFrm001Location"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001Agent = applyTask.Task.CurrentDocument.Fields["HREngFrm001Agent"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001Transp = applyTask.Task.CurrentDocument.Fields["HREngFrm001Transp"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001LicPlate = applyTask.Task.CurrentDocument.Fields["HREngFrm001LicPlate"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001Cause = applyTask.Task.CurrentDocument.Fields["HREngFrm001Cause"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001DefOutTime = applyTask.Task.CurrentDocument.Fields["HREngFrm001DefOutTime"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001FF = applyTask.Task.CurrentDocument.Fields["HREngFrm001FF"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001OutTime = applyTask.Task.CurrentDocument.Fields["HREngFrm001OutTime"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001DefBakTime = applyTask.Task.CurrentDocument.Fields["HREngFrm001DefBakTime"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001CH = applyTask.Task.CurrentDocument.Fields["HREngFrm001CH"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001BakTime = applyTask.Task.CurrentDocument.Fields["HREngFrm001BakTime"].FieldValue.ToString().Trim();

            if(HREngFrm001.HREngFrm001PIR.Equals("否"))
            {
                string account = HREngFrm001.HREngFrm001User;
                account = account.Substring(4, 6);
              
                HREngFrm001.CRADNO = SEARCHCARDNO(account);
            }
            else if (HREngFrm001.HREngFrm001PIR.Equals("是"))
            {
                string account = HREngFrm001.HREngFrm001CN;
                HREngFrm001.CRADNO = SEARCHCARDNO(account);

                HREngFrm001.HREngFrm001User = HREngFrm001.HREngFrm001OU+ HREngFrm001.HREngFrm001CN;
                
            }
           



            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                if (!string.IsNullOrEmpty(HREngFrm001.TaskId))
                {
                    ADDTKGAFFAIRSHREngFrm001(HREngFrm001);
                }
            }

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public void ADDTKGAFFAIRSHREngFrm001(DATAHREngFrm001 HREngFrm001)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            //queryString.AppendFormat(@" INSERT INTO [TK].dbo.COPMA");
            //queryString.AppendFormat(@" (COMPANY,MA001,MA002)");
            //queryString.AppendFormat(@" VALUES (@MA001,@MA001,@MA002)");


            queryString.AppendFormat(@" 
                                     INSERT INTO  [TKGAFFAIRS].[dbo].[HREngFrm001]
                                     ([TaskId],[HREngFrm001SN],[HREngFrm001Date],[HREngFrm001PIR],[HREngFrm001User],[HREngFrm001UsrDpt],[HREngFrm001Rank],[HREngFrm001OutDate],[HREngFrm001Location],[HREngFrm001Agent],[HREngFrm001Transp],[HREngFrm001LicPlate],[HREngFrm001Cause],[HREngFrm001DefOutTime],[HREngFrm001FF],[HREngFrm001OutTime],[HREngFrm001DefBakTime],[HREngFrm001CH],[HREngFrm001BakTime],[CRADNO])
                                     VALUES
                                     (@TaskId,@HREngFrm001SN,@HREngFrm001Date,@HREngFrm001PIR,@HREngFrm001User,@HREngFrm001UsrDpt,@HREngFrm001Rank,@HREngFrm001OutDate,@HREngFrm001Location,@HREngFrm001Agent,@HREngFrm001Transp,@HREngFrm001LicPlate,@HREngFrm001Cause,@HREngFrm001DefOutTime,@HREngFrm001FF,@HREngFrm001OutTime,@HREngFrm001DefBakTime,@HREngFrm001CH,@HREngFrm001BakTime,@CRADNO)
    
                                     ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TaskId", SqlDbType.NVarChar).Value = HREngFrm001.TaskId;
                    command.Parameters.Add("@HREngFrm001SN", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001SN;
                    command.Parameters.Add("@HREngFrm001Date", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001Date;
                    command.Parameters.Add("@HREngFrm001PIR", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001PIR;
                    command.Parameters.Add("@HREngFrm001User", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001User;
                    command.Parameters.Add("@HREngFrm001UsrDpt", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001UsrDpt;
                    command.Parameters.Add("@HREngFrm001Rank", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001Rank;
                    command.Parameters.Add("@HREngFrm001OutDate", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001OutDate;
                    command.Parameters.Add("@HREngFrm001Location", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001Location;
                    command.Parameters.Add("@HREngFrm001Agent", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001Agent;
                    command.Parameters.Add("@HREngFrm001Transp", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001Transp;
                    command.Parameters.Add("@HREngFrm001LicPlate", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001LicPlate;
                    command.Parameters.Add("@HREngFrm001Cause", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001Cause;
                    command.Parameters.Add("@HREngFrm001DefOutTime", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001DefOutTime;
                    command.Parameters.Add("@HREngFrm001FF", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001FF;
                    command.Parameters.Add("@HREngFrm001OutTime", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001OutTime;
                    command.Parameters.Add("@HREngFrm001DefBakTime", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001DefBakTime;
                    command.Parameters.Add("@HREngFrm001CH", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001CH;
                    command.Parameters.Add("@HREngFrm001BakTime", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001BakTime;
                    command.Parameters.Add("@CRADNO", SqlDbType.NVarChar).Value = HREngFrm001.CRADNO;



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

        public string SEARCHCARDNO(string Account)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["CHIYUconnectionstring"].ToString();
            Ede.Uof.Utility.Data.DatabaseHelper m_db = new Ede.Uof.Utility.Data.DatabaseHelper(connectionString);

            string cmdTxt = @" SELECT [EmployeeID],[CardNo]
                                FROM [CHIYU].[dbo].[Person]
                                WHERE [EmployeeID]=@Account
                        ";

            m_db.AddParameter("@Account", Account);
           

            DataTable dt = new DataTable();

            dt.Load(m_db.ExecuteReader(cmdTxt));

            if(dt.Rows.Count>0)
            {
                return dt.Rows[0]["CardNo"].ToString();
            }
            else
            {
                return null;
            }
        }
    }
}
