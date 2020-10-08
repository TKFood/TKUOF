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

namespace TKUOF.TRIGGER.GAFrm
{
    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public class DATAUOFGAFrm
        {
            public string TaskId;
            public string GAFrm004SN;
            public string SERNO;
            public string GAFrm004SI;
            public string GAFrm004CM;
            public string GAFrm004OD;
            public string [] GAFrm004DN = new string[99];
            public string [] GAFrm004NB=new string[99];
            public string [] GAFrm004ID = new string[99];
            public string [] GAFrm004ER = new string[99];
            public string [] GAFrm004S0ND = new string[99];
            public string GAFrm004PS;
            public string GAFrm004PID;
            public string GAFrm004RD;



        }
        public void Finally()
        {
           
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            DATAUOFGAFrm UOFGAFrm = new DATAUOFGAFrm();
  

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            UOFGAFrm.TaskId = applyTask.Task.TaskId;

            UOFGAFrm.GAFrm004SN = applyTask.Task.CurrentDocument.Fields["GAFrm004SN"].FieldValue.ToString().Trim();
            UOFGAFrm.GAFrm004SI = applyTask.Task.CurrentDocument.Fields["GAFrm004SI"].FieldValue.ToString().Trim();
            UOFGAFrm.GAFrm004CM = applyTask.Task.CurrentDocument.Fields["GAFrm004CM"].FieldValue.ToString().Trim();
            UOFGAFrm.GAFrm004OD = applyTask.Task.CurrentDocument.Fields["GAFrm004OD"].FieldValue.ToString().Trim();
            UOFGAFrm.GAFrm004PS = applyTask.Task.CurrentDocument.Fields["GAFrm004PS"].FieldValue.ToString().Trim();
            UOFGAFrm.GAFrm004PID = applyTask.Task.CurrentDocument.Fields["GAFrm004PID"].FieldValue.ToString().Trim();
            UOFGAFrm.GAFrm004RD = applyTask.Task.CurrentDocument.Fields["GAFrm004RD"].FieldValue.ToString().Trim();

            XmlNodeList NodeLists = applyTask.Task.CurrentDocument.;


            foreach (XmlNode OneNode in NodeLists)
            {
                String StrAttrName = OneNode.Attributes.Name;
                String StrAttrValue = OneNode.Attributes[" MyAttr1 "].Value;
                String StrAttrValue = OneNode.InnerText;
            }



            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                if (!string.IsNullOrEmpty(UOFGAFrm.TaskId))
                {
                    ADDTKGAFFAIRSUOFGAFrm(UOFGAFrm);
                }
            }

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public void ADDTKGAFFAIRSUOFGAFrm(DATAUOFGAFrm UOFGAFrm)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            //queryString.AppendFormat(@" INSERT INTO [TK].dbo.COPMA");
            //queryString.AppendFormat(@" (COMPANY,MA001,MA002)");
            //queryString.AppendFormat(@" VALUES (@MA001,@MA001,@MA002)");

            queryString.AppendFormat(@" INSERT INTO  [TKGAFFAIRS].[dbo].[HREngFrm001]");
            queryString.AppendFormat(@" ([TaskId],[HREngFrm001SN],[HREngFrm001Date],[HREngFrm001PIR],[HREngFrm001User],[HREngFrm001UsrDpt],[HREngFrm001Rank],[HREngFrm001OutDate],[HREngFrm001Location],[HREngFrm001Agent],[HREngFrm001Transp],[HREngFrm001LicPlate],[HREngFrm001Cause],[HREngFrm001DefOutTime],[HREngFrm001FF],[HREngFrm001OutTime],[HREngFrm001DefBakTime],[HREngFrm001CH],[HREngFrm001BakTime],[CRADNO])");
            queryString.AppendFormat(@" VALUES");
            queryString.AppendFormat(@" (@TaskId,@HREngFrm001SN,@HREngFrm001Date,@HREngFrm001PIR,@HREngFrm001User,@HREngFrm001UsrDpt,@HREngFrm001Rank,@HREngFrm001OutDate,@HREngFrm001Location,@HREngFrm001Agent,@HREngFrm001Transp,@HREngFrm001LicPlate,@HREngFrm001Cause,@HREngFrm001DefOutTime,@HREngFrm001FF,@HREngFrm001OutTime,@HREngFrm001DefBakTime,@HREngFrm001CH,@HREngFrm001BakTime,@CRADNO)");
            queryString.AppendFormat(@" ");
            queryString.AppendFormat(@" ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TaskId", SqlDbType.NVarChar).Value = HREngFrm001.TaskId;
                   



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
