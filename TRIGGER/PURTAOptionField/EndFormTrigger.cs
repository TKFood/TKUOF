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
using Ede.Uof.WKF.CustomExternal;

namespace TKUOF.TRIGGER.PURTAOptionField
{
    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public class DATAPURTA
        {
            public string TaskId;

            public string CREATOR;
            public string USR_GROUP;
            public string CREATE_DATE;
            public string MODIFIER;
            public string MODI_DATE;
            public string FLAG;
            public string CREATE_TIME;
            public string MODI_TIME;
            public string TRANS_TYPE;
            public string TRANS_NAME;
            public string sync_date;
            public string sync_time;
            public string sync_mark;
            public string sync_count;
            public string DataUser;
            public string DataGroup;

            public string TA001;
            public string TA002;
            public string TA003;
            public string TA004;
            public string TA005;
            public string TA006;
            public string TA007;
            public string TA008;
            public string TA009;
            public string TA010;
            public string TA011;
            public string TA012;
            public string TA013;
            public string TA014;
            public string TA015;
            public string TA016;
            public string TA017;
            public string TA018;
            public string TA019;
            public string TA020;
            public string TA021;
            public string TA022;
            public string TA023;
            public string TA024;
            public string TA025;
            public string TA026;
            public string TA027;
            public string TA028;
            public string TA029;
            public string TA030;
            public string TA031;
            public string TA032;
            public string TA033;
            public string TA034;
            public string TA035;
            public string TA036;
            public string TA037;
            public string TA038;
            public string TA039;
            public string TA040;
            public string TA041;
            public string TA042;
            public string TA043;
            public string TA044;
            public string TA045;
            public string TA046;
            public string UDF01;
            public string UDF02;
            public string UDF03;
            public string UDF04;
            public string UDF05;
            public string UDF06;
            public string UDF07;
            public string UDF08;
            public string UDF09;
            public string UDF10;

        }

        public class DATAPURTB
        {
            public string TaskId;

            public string CREATOR;
            public string USR_GROUP;
            public string CREATE_DATE;
            public string MODIFIER;
            public string MODI_DATE;
            public string FLAG;
            public string CREATE_TIME;
            public string MODI_TIME;
            public string TRANS_TYPE;
            public string TRANS_NAME;
            public string sync_date;
            public string sync_time;
            public string sync_mark;
            public string sync_count;
            public string DataUser;
            public string DataGroup;
            public string TB001;
            public string TB002;
            public string TB003;
            public string TB004;
            public string TB005;
            public string TB006;
            public string TB007;
            public string TB008;
            public string TB009;
            public string TB010;
            public string TB011;
            public string TB012;
            public string TB013;
            public string TB014;
            public string TB015;
            public string TB016;
            public string TB017;
            public string TB018;
            public string TB019;
            public string TB020;
            public string TB021;
            public string TB022;
            public string TB023;
            public string TB024;
            public string TB025;
            public string TB026;
            public string TB027;
            public string TB028;
            public string TB029;
            public string TB030;
            public string TB031;
            public string TB032;
            public string TB033;
            public string TB034;
            public string TB035;
            public string TB036;
            public string TB037;
            public string TB038;
            public string TB039;
            public string TB040;
            public string TB041;
            public string TB042;
            public string TB043;
            public string TB044;
            public string TB045;
            public string TB046;
            public string TB047;
            public string TB048;
            public string TB049;
            public string TB050;
            public string TB051;
            public string TB052;
            public string TB053;
            public string TB054;
            public string TB055;
            public string TB056;
            public string TB057;
            public string TB058;
            public string TB059;
            public string TB060;
            public string TB061;
            public string TB062;
            public string TB063;
            public string TB064;
            public string TB065;
            public string TB066;
            public string TB067;
            public string TB068;
            public string TB069;
            public string TB070;
            public string TB071;
            public string TB072;
            public string TB073;
            public string TB074;
            public string TB075;
            public string TB076;
            public string TB077;
            public string TB078;
            public string TB079;
            public string TB080;
            public string TB081;
            public string TB082;
            public string TB083;
            public string TB084;
            public string TB085;
            public string TB086;
            public string TB087;
            public string TB088;
            public string TB089;
            public string TB090;
            public string TB091;
            public string TB092;
            public string TB093;
            public string TB094;
            public string TB095;
            public string TB096;
            public string TB097;
            public string TB098;
            public string TB099;
            public string UDF01;
            public string UDF02;
            public string UDF03;
            public string UDF04;
            public string UDF05;
            public string UDF06;
            public string UDF07;
            public string UDF08;
            public string UDF09;
            public string UDF10;

        }


        public void Finally()
        {
            
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            int ROWS = 1;
            StringBuilder PURTBSB = new StringBuilder();

            DATAPURTA DataPURTA = new DATAPURTA();
            DATAPURTB DataPURTB = new DATAPURTB();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);

            //MA001 = applyTask.Task.CurrentDocument.Fields["MA001"].FieldValue.ToString().Trim();
            //MA002 = applyTask.Task.CurrentDocument.Fields["MA002"].FieldValue.ToString().Trim();

            //針對主檔抓出來的資料作處理
            XmlNode node = xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='PURTAB']");

            DataPURTA.TaskId = applyTask.Task.TaskId;
            DataPURTA.TA001 = "A311";
            DataPURTA.TA002 = FINDMAXPURTATA002("A311");

            DataPURTA.TA004 = node.SelectSingleNode("FieldValue").Attributes["DEP"].Value;
            DataPURTA.TA006 = node.SelectSingleNode("FieldValue").Attributes["COMMENT"].Value;
            DataPURTA.TA012 = node.SelectSingleNode("FieldValue").Attributes["NAME"].Value;

            //針對DETAIL抓出來的資料作處理

            foreach (XmlNode nodeDetail in xmlDoc.SelectNodes("./Form/FormFieldValue/FieldItem[@fieldId='PURTAB']/FieldValue/Item"))
            {
                DataPURTB.TB001 = "A311";
                DataPURTB.TB002 = DataPURTA.TA002;
                DataPURTB.TB003 = ROWS.ToString().PadLeft(4, '0');
                DataPURTB.TB004 = nodeDetail.Attributes["品號"].Value;

                PURTBSB.Append(SETADDPURDTB(DataPURTB));
                PURTBSB.AppendLine();

                ROWS = ROWS + 1;

            }

            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                if (!string.IsNullOrEmpty(DataPURTA.TA001) && !string.IsNullOrEmpty(DataPURTA.TA002))
                {
                    UPDATETB_WKF_TASK(applyTask, DataPURTA.TA001, DataPURTA.TA002);
                    INSERTINTOPURTAB(DataPURTA, PURTBSB, applyTask.FormNumber);
                }
            }

            return "";
        }

        public void OnError(Exception errorException)
        {
          
        }

        public void UPDATETB_WKF_TASK(ApplyTask applyTask,string TA001,string TA002)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);

            XmlElement att = (XmlElement)xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='PURTAB']/FieldValue");
            att.SetAttribute("TA001", TA001);
            att.SetAttribute("TA002", TA002);
            //string str = xmlDoc.OuterXml;

            string connectionString = ConfigurationManager.ConnectionStrings["connectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@" 
                                        UPDATE [UOFTEST].dbo.TB_WKF_TASK
                                        SET CURRENT_DOC=@CURRENT_DOC
                                        WHERE TASK_ID=@TASK_ID
                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@CURRENT_DOC", SqlDbType.NVarChar).Value = xmlDoc.OuterXml;
                    command.Parameters.Add("@TASK_ID", SqlDbType.NVarChar).Value = applyTask.TaskId;

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

        public string FINDMAXPURTATA002(string TA001)
        {
            String connectionString;
            SqlConnection conn;
            connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ConnectionString;
            conn = new SqlConnection(connectionString);
            DataSet PURTADS = new DataSet();
            string TA002 = "";      
          
            //找出請購最大的單號
            StringBuilder sbSql = new StringBuilder();
            sbSql.AppendFormat(@" 
                                SELECT ISNULL(MAX(TA002),'00000000000') AS TA002
                                FROM [TK].[dbo].[PURTA] 
                                WHERE  TA001='{0}' AND TA003='{1}'
                                ", TA001, DateTime.Now.ToString("yyyyMMdd"));

          
            SqlCommand cmd = new SqlCommand(sbSql.ToString(), conn);           
            conn.Open();
            SqlDataAdapter adapt = new SqlDataAdapter(cmd);
            adapt.Fill(PURTADS);


            TA002 = SETTA002(PURTADS.Tables[0].Rows[0]["TA002"].ToString());
            return TA002;
        }

        public string SETTA002(string TA002)
        {
            if (TA002.Equals("00000000000"))
            {
                return DateTime.Now.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(TA002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return DateTime.Now.ToString("yyyyMMdd") + temp.ToString();
            }
        }

        public StringBuilder SETADDPURDTB(DATAPURTB DataPURTB)
        {
            StringBuilder ADDPURDTB = new StringBuilder();
            ADDPURDTB.AppendFormat(@"
                                    INSERT INTO [TK].dbo.PURTB
                                    (TB001,TB002,TB003,TB004)
                                    VALUES
                                    ('{0}','{1}','{2}','{3}')
                                    ", DataPURTB.TB001, DataPURTB.TB002, DataPURTB.TB003, DataPURTB.TB004);

            return ADDPURDTB;
        }

        public void INSERTINTOPURTAB(DATAPURTA DataPURTA, StringBuilder PURTBSB,string FormNumber)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                       INSERT INTO [TK].dbo.PURTA
                                        (TA001,TA002,TA004,TA005,TA006,TA012)
                                        VALUES
                                        (@TA001,@TA002,@TA004,@TA005,@TA006,@TA012) 
                                    ");
            queryString.AppendLine();
            queryString.Append(PURTBSB.ToString());

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TA001", SqlDbType.NVarChar).Value = DataPURTA.TA001;
                    command.Parameters.Add("@TA002", SqlDbType.NVarChar).Value = DataPURTA.TA002;
                    command.Parameters.Add("@TA004", SqlDbType.NVarChar).Value = DataPURTA.TA004;
                    command.Parameters.Add("@TA005", SqlDbType.NVarChar).Value = FormNumber;
                    command.Parameters.Add("@TA006", SqlDbType.NVarChar).Value = DataPURTA.TA006;
                    command.Parameters.Add("@TA012", SqlDbType.NVarChar).Value = DataPURTA.TA012;

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

