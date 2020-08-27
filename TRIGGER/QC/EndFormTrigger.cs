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

namespace TKUOF.TRIGGER.QC
{
    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public class DATAQC
        {
            public string TaskId;

            public string QCFrm002SN;
            public string QCFrm002QCC;
            public string QCFrm002PN;



        }
        public void Finally()
        {
           
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            DATAQC QC = new DATAQC();
           
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            QC.TaskId = applyTask.Task.TaskId;

            QC.QCFrm002SN = applyTask.Task.CurrentDocument.Fields["QCFrm002SN"].FieldValue.ToString().Trim();
            QC.QCFrm002QCC = applyTask.Task.CurrentDocument.Fields["QCFrm002QCC"].FieldValue.ToString().Trim();
            QC.QCFrm002PN = applyTask.Task.CurrentDocument.Fields["QCFrm002PN"].FieldValue.ToString().Trim();


            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                if (!string.IsNullOrEmpty(QC.TaskId))
                {
                    ADDTBFORMQC(QC);
                }
            }

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public void ADDTBFORMQC(DATAQC QC)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            //queryString.AppendFormat(@" INSERT INTO [TK].dbo.COPMA");
            //queryString.AppendFormat(@" (COMPANY,MA001,MA002)");
            //queryString.AppendFormat(@" VALUES (@MA001,@MA001,@MA002)");

           
            queryString.AppendFormat(@"
                                    INSERT INTO [TKQC].[dbo].[TBFORMQC]
                                    ([ID],[TaskId],[QCFrm002SN],[QCFrm002QCC],[QCFrm002PN])
                                    VALUES
                                    (@ID,@TaskId,@QCFrm002SN,@QCFrm002QCC,@QCFrm002PN)
                                    ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@ID", SqlDbType.UniqueIdentifier).Value = Guid.NewGuid();
                    command.Parameters.Add("@TaskId", SqlDbType.NVarChar).Value = QC.TaskId;
                    command.Parameters.Add("@QCFrm002SN", SqlDbType.NVarChar).Value = QC.QCFrm002SN;
                    command.Parameters.Add("@QCFrm002QCC", SqlDbType.NVarChar).Value = QC.QCFrm002QCC;
                    command.Parameters.Add("@QCFrm002PN", SqlDbType.NVarChar).Value = QC.QCFrm002PN;


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
