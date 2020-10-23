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
        int ROWS = 0;

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



            //針對DETAIL抓出來的資料作處理

            foreach (XmlNode node in xmlDoc.SelectNodes("./Form/FormFieldValue/FieldItem[@fieldId='GAFrm004IT']/DataGrid/Row"))
            {
                UOFGAFrm.GAFrm004DN[ROWS] = node.SelectSingleNode("./Cell[@fieldId='GAFrm004DN']").Attributes["fieldValue"].Value;
                UOFGAFrm.GAFrm004NB[ROWS] = node.SelectSingleNode("./Cell[@fieldId='GAFrm004NB']").Attributes["fieldValue"].Value;
                UOFGAFrm.GAFrm004ID[ROWS] = node.SelectSingleNode("./Cell[@fieldId='GAFrm004ID']").Attributes["fieldValue"].Value;
                UOFGAFrm.GAFrm004ER[ROWS] = node.SelectSingleNode("./Cell[@fieldId='GAFrm004ER']").Attributes["fieldValue"].Value;
                UOFGAFrm.GAFrm004S0ND[ROWS] = node.SelectSingleNode("./Cell[@fieldId='GAFrm004S0ND']").Attributes["fieldValue"].Value;

                ROWS = ROWS + 1;
            }


            ///1a31c995-f2e1-40cc-9cb9-6079aca2a242 副總 正式
            ///3adb6f7a-98b5-42e5-8dfb-21a9e3f680ae 葉志剛 正式
            ///0077c97a-8699-4688-be7e-ea1ecb960145 葉志剛 test

            //表單核準後
            //if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)

            //指定某人
            if (applyTask.Task.CurrentSite.CurrentNode.ActualSignerId == "1a31c995-f2e1-40cc-9cb9-6079aca2a242" && applyTask.Task.CurrentSite.SiteResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
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
            int SERNO = 0;

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@" ");

            try
            {

                queryString.AppendFormat(@" 
                                        INSERT INTO [TKGAFFAIRS].[dbo].[UOFGAFrm]
                                        ([TaskId],[GAFrm004SN],[SERNO],[GAFrm004SI],[GAFrm004CM],[GAFrm004OD],[GAFrm004DN],[GAFrm004NB],[GAFrm004ID],[GAFrm004ER],[GAFrm004S0ND],[GAFrm004PS],[GAFrm004PID],[GAFrm004RD])
                                        VALUES
                                        (@TaskId,@GAFrm004SN,@SERNO,@GAFrm004SI,@GAFrm004CM,@GAFrm004OD,@GAFrm004DN,@GAFrm004NB,@GAFrm004ID,@GAFrm004ER,@GAFrm004S0ND,@GAFrm004PS,@GAFrm004PID,@GAFrm004RD)
                                        ");


                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);

                    command.Connection.Open();

                    //command.Parameters.Clear();//清除掉目前宣告出來的Parameters
                    //command.Parameters.Add("@TaskId", SqlDbType.NVarChar).Value = UOFGAFrm.TaskId;
                    //command.Parameters.Add("@GAFrm004SN", SqlDbType.NVarChar).Value = UOFGAFrm.GAFrm004SN;
                    //command.Parameters.Add("@SERNO", SqlDbType.NVarChar).Value = SERNO.ToString();
                    //int count = command.ExecuteNonQuery();

                    for (int i = 0; i < ROWS; i++)
                    {

                        command.Parameters.Clear();//清除掉目前宣告出來的Parameters
                        command.Parameters.Add("@TaskId", SqlDbType.NVarChar).Value = UOFGAFrm.TaskId;
                        command.Parameters.Add("@GAFrm004SN", SqlDbType.NVarChar).Value = UOFGAFrm.GAFrm004SN;

                        SERNO = i + 1;
                        command.Parameters.Add("@SERNO", SqlDbType.NVarChar).Value = SERNO.ToString();
                        command.Parameters.Add("@GAFrm004SI", SqlDbType.NVarChar).Value = UOFGAFrm.GAFrm004SI;
                        command.Parameters.Add("@GAFrm004CM", SqlDbType.NVarChar).Value = UOFGAFrm.GAFrm004CM;
                        command.Parameters.Add("@GAFrm004OD", SqlDbType.NVarChar).Value = UOFGAFrm.GAFrm004OD;
                        command.Parameters.Add("@GAFrm004DN", SqlDbType.NVarChar).Value = UOFGAFrm.GAFrm004DN[i].ToString();
                        command.Parameters.Add("@GAFrm004NB", SqlDbType.NVarChar).Value = UOFGAFrm.GAFrm004NB[i].ToString();
                        command.Parameters.Add("@GAFrm004ID", SqlDbType.NVarChar).Value = UOFGAFrm.GAFrm004ID[i].ToString();
                        command.Parameters.Add("@GAFrm004ER", SqlDbType.NVarChar).Value = UOFGAFrm.GAFrm004ER[i].ToString();
                        command.Parameters.Add("@GAFrm004S0ND", SqlDbType.NVarChar).Value = UOFGAFrm.GAFrm004S0ND[i].ToString();
                        command.Parameters.Add("@GAFrm004PS", SqlDbType.NVarChar).Value = UOFGAFrm.GAFrm004PS;
                        command.Parameters.Add("@GAFrm004PID", SqlDbType.NVarChar).Value = UOFGAFrm.GAFrm004PID;
                        command.Parameters.Add("@GAFrm004RD", SqlDbType.NVarChar).Value = UOFGAFrm.GAFrm004RD;


                        int count = command.ExecuteNonQuery();


                    }

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
