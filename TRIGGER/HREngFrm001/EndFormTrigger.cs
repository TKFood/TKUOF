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
            public string HREngFrm001SN;
            public string HREngFrm001UsrDpt;
            public string HREngFrm001User;
            public string HREngFrm001Name;
            public string HREngFrm001Dpt;
            public string HREngFrm001Date;
            public string HREngFrm001TITLE;
            public string HREngFrm001Agent;
            public string HREngFrm001Transp;
            public string HREngFrm001Location;
            public string HREngFrm001Cause;
            public string HREngFrm001DefOutTime;
            public string HREngFrm001OutTime;
            public string HREngFrm001DefBakTime;
            public string HREngFrm001BakTime;
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
            HREngFrm001.HREngFrm001SN = applyTask.Task.CurrentDocument.Fields["HREngFrm001SN"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001UsrDpt = applyTask.Task.CurrentDocument.Fields["HREngFrm001SN"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001User = applyTask.Task.CurrentDocument.Fields["HREngFrm001User"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001Name = applyTask.Task.CurrentDocument.Fields["HREngFrm001Name"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001Dpt = applyTask.Task.CurrentDocument.Fields["HREngFrm001Dpt"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001Date = applyTask.Task.CurrentDocument.Fields["HREngFrm001Date"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001TITLE = applyTask.Task.CurrentDocument.Fields["HREngFrm001TITLE"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001Agent = applyTask.Task.CurrentDocument.Fields["HREngFrm001Agent"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001Transp = applyTask.Task.CurrentDocument.Fields["HREngFrm001Transp"].FieldValue.ToString().Trim(); 
            HREngFrm001.HREngFrm001Location = applyTask.Task.CurrentDocument.Fields["HREngFrm001Location"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001Cause = applyTask.Task.CurrentDocument.Fields["HREngFrm001Cause"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001DefOutTime = applyTask.Task.CurrentDocument.Fields["HREngFrm001DefOutTime"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001OutTime = applyTask.Task.CurrentDocument.Fields["HREngFrm001OutTime"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001DefBakTime = applyTask.Task.CurrentDocument.Fields["HREngFrm001DefBakTime"].FieldValue.ToString().Trim();
            HREngFrm001.HREngFrm001BakTime = applyTask.Task.CurrentDocument.Fields["HREngFrm001BakTime"].FieldValue.ToString().Trim(); ;

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

            queryString.AppendFormat(@" INSERT INTO  [TKGAFFAIRS].[dbo].[HREngFrm001]");
            queryString.AppendFormat(@" ([TaskId],[HREngFrm001SN],[HREngFrm001UsrDpt],[HREngFrm001User],[HREngFrm001Name],[HREngFrm001Dpt],[HREngFrm001Date],[HREngFrm001TITLE],[HREngFrm001Agent],[HREngFrm001Transp],[HREngFrm001Location],[HREngFrm001Cause],[HREngFrm001DefOutTime],[HREngFrm001OutTime],[HREngFrm001DefBakTime],[HREngFrm001BakTime])");
            queryString.AppendFormat(@" VALUES");
            queryString.AppendFormat(@" (@TaskId,@HREngFrm001SN,@HREngFrm001UsrDpt,@HREngFrm001User,@HREngFrm001Name,@HREngFrm001Dpt,@HREngFrm001Date,@HREngFrm001TITLE,@HREngFrm001Agent,@HREngFrm001Transp,@HREngFrm001Location,@HREngFrm001Cause,@HREngFrm001DefOutTime,@HREngFrm001OutTime,@HREngFrm001DefBakTime,@HREngFrm001BakTime)");
            queryString.AppendFormat(@" ");
            queryString.AppendFormat(@" ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TaskId", SqlDbType.NVarChar).Value = HREngFrm001.TaskId;
                    command.Parameters.Add("@HREngFrm001SN", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001SN;
                    command.Parameters.Add("@HREngFrm001UsrDpt", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001UsrDpt;
                    command.Parameters.Add("@HREngFrm001User", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001User;
                    command.Parameters.Add("@HREngFrm001Name", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001Name;
                    command.Parameters.Add("@HREngFrm001Dpt", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001Dpt;
                    command.Parameters.Add("@HREngFrm001Date", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001Date;
                    command.Parameters.Add("@HREngFrm001TITLE", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001TITLE;
                    command.Parameters.Add("@HREngFrm001Agent", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001Agent;
                    command.Parameters.Add("@HREngFrm001Transp", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001Transp;
                    command.Parameters.Add("@HREngFrm001Location", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001Location;
                    command.Parameters.Add("@HREngFrm001Cause", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001Cause;
                    command.Parameters.Add("@HREngFrm001DefOutTime", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001DefOutTime;
                    command.Parameters.Add("@HREngFrm001OutTime", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001OutTime;
                    command.Parameters.Add("@HREngFrm001DefBakTime", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001DefBakTime;
                    command.Parameters.Add("@HREngFrm001BakTime", SqlDbType.NVarChar).Value = HREngFrm001.HREngFrm001BakTime;

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
