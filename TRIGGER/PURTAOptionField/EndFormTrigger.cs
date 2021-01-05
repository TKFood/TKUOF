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
        UserUCO UserUCOSuperior = new UserUCO();
        string account;
        string userGuid;


        public class DATAPURTA
        {
            public string TaskId;

            public string COMPANY;
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

            public string COMPANY;
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

            //找出最後的簽核者
            account = xmlDoc.SelectSingleNode("./Form/Applicant").Attributes["account"].Value;

            //MA001 = applyTask.Task.CurrentDocument.Fields["MA001"].FieldValue.ToString().Trim();
            //MA002 = applyTask.Task.CurrentDocument.Fields["MA002"].FieldValue.ToString().Trim();

            //針對主檔抓出來的資料作處理
            XmlNode node = xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='PURTAB']");

            DataPURTA.TaskId = applyTask.Task.TaskId;

            DataPURTA.COMPANY = "TK";
            DataPURTA.CREATOR = node.SelectSingleNode("FieldValue").Attributes["NAME"].Value;
            DataPURTA.USR_GROUP = node.SelectSingleNode("FieldValue").Attributes["DEP"].Value;
            DataPURTA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
            DataPURTA.MODIFIER = account;
            DataPURTA.MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            DataPURTA.FLAG = "1";
            DataPURTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
            DataPURTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
            DataPURTA.TRANS_TYPE = "P001";
            DataPURTA.TRANS_NAME = "PURI05";
            DataPURTA.sync_date = "";
            DataPURTA.sync_time = "";
            DataPURTA.sync_mark = "";
            DataPURTA.sync_count = "0";
            DataPURTA.DataUser = "";
            DataPURTA.DataGroup = node.SelectSingleNode("FieldValue").Attributes["DEP"].Value;

            DataPURTA.TA001 = "A311";
            DataPURTA.TA002 = FINDMAXPURTATA002("A311");
            DataPURTA.TA003 = DateTime.Now.ToString("yyyyMMdd");
            DataPURTA.TA004 = node.SelectSingleNode("FieldValue").Attributes["DEP"].Value;
            DataPURTA.TA005 = applyTask.Task.TaskId;
            DataPURTA.TA006 = node.SelectSingleNode("FieldValue").Attributes["COMMENT"].Value;
            DataPURTA.TA007 = "N";
            DataPURTA.TA008 = "0";
            DataPURTA.TA009 = "9";
            DataPURTA.TA010 = "20";
            DataPURTA.TA011 = "0";
            DataPURTA.TA012 = node.SelectSingleNode("FieldValue").Attributes["NAME"].Value;
            DataPURTA.TA013 = DateTime.Now.ToString("yyyyMMdd");
            DataPURTA.TA014 = account;
            DataPURTA.TA015 = "0";
            DataPURTA.TA016 = "N";
            DataPURTA.TA017 = "0";
            DataPURTA.TA018 = "";
            DataPURTA.TA019 = "";
            DataPURTA.TA020 = "0";
            DataPURTA.TA021 = "";
            DataPURTA.TA022 = "";
            DataPURTA.TA023 = "0";
            DataPURTA.TA024 = "0";
            DataPURTA.TA025 = "";
            DataPURTA.TA026 = "";
            DataPURTA.TA027 = "";
            DataPURTA.TA028 = "";
            DataPURTA.TA029 = "";
            DataPURTA.TA030 = "0";
            DataPURTA.TA031 = "";
            DataPURTA.TA032 = "0";
            DataPURTA.TA033 = "";
            DataPURTA.TA034 = "";
            DataPURTA.TA035 = "";
            DataPURTA.TA036 = "0";
            DataPURTA.TA037 = "0";
            DataPURTA.TA038 = "0";
            DataPURTA.TA039 = "0";
            DataPURTA.TA040 = "0";
            DataPURTA.TA041 = "";
            DataPURTA.TA042 = "";
            DataPURTA.TA043 = "";
            DataPURTA.TA044 = "";
            DataPURTA.TA045 = "";
            DataPURTA.TA046 = "";
            DataPURTA.UDF01 = "";
            DataPURTA.UDF02 = "";
            DataPURTA.UDF03 = "";
            DataPURTA.UDF04 = "";
            DataPURTA.UDF05 = "";
            DataPURTA.UDF06 = "0";
            DataPURTA.UDF07 = "0";
            DataPURTA.UDF08 = "0";
            DataPURTA.UDF09 = "0";
            DataPURTA.UDF10 = "0";


            //針對DETAIL抓出來的資料作處理

            foreach (XmlNode nodeDetail in xmlDoc.SelectNodes("./Form/FormFieldValue/FieldItem[@fieldId='PURTAB']/FieldValue/Item"))
            {
                DataPURTB.COMPANY = DataPURTA.COMPANY;
                DataPURTB.CREATOR = DataPURTA.CREATOR;
                DataPURTB.USR_GROUP = DataPURTA.USR_GROUP;
                DataPURTB.CREATE_DATE = DataPURTA.CREATE_DATE;
                DataPURTB.MODIFIER = DataPURTA.MODIFIER;
                DataPURTB.MODI_DATE = DataPURTA.MODI_DATE;
                DataPURTB.FLAG = DataPURTA.FLAG;
                DataPURTB.CREATE_TIME = DataPURTA.CREATE_TIME;
                DataPURTB.MODI_TIME = DataPURTA.MODI_TIME;
                DataPURTB.TRANS_TYPE = DataPURTA.TRANS_TYPE;
                DataPURTB.TRANS_NAME = DataPURTA.TRANS_NAME;
                DataPURTB.sync_date = DataPURTA.sync_date;
                DataPURTB.sync_time = DataPURTA.sync_time;
                DataPURTB.sync_mark = DataPURTA.sync_mark;
                DataPURTB.sync_count = DataPURTA.sync_count;
                DataPURTB.DataUser = DataPURTA.DataUser;
                DataPURTB.DataGroup = DataPURTA.DataGroup;

                DataPURTB.TB001 = "A311";
                DataPURTB.TB002 = DataPURTA.TA002;
                DataPURTB.TB003 = ROWS.ToString().PadLeft(4, '0');
                DataPURTB.TB004 = nodeDetail.Attributes["品號"].Value;
                DataPURTB.TB009 = nodeDetail.Attributes["數量"].Value;
                DataPURTB.TB011 = nodeDetail.Attributes["需求日"].Value;
                DataPURTB.TB012 = nodeDetail.Attributes["單身備註"].Value;
                DataPURTB.TB024 = nodeDetail.Attributes["單身備註"].Value;

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

        public string FINDCMSMVMV004(string MV001)
        {
            String connectionString;
            SqlConnection conn;
            connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ConnectionString;
            conn = new SqlConnection(connectionString);
            DataSet PURTADS = new DataSet();
            string MV004 = "";

            //找出請購最大的單號
            StringBuilder sbSql = new StringBuilder();
            sbSql.AppendFormat(@"                                 
                                SELECT MV004
                                FROM [TK].dbo.CMSMV
                                WHERE MV001='{0}'
                                ", MV001);


            SqlCommand cmd = new SqlCommand(sbSql.ToString(), conn);
            conn.Open();
            SqlDataAdapter adapt = new SqlDataAdapter(cmd);
            adapt.Fill(PURTADS);


            MV004 = SETTA002(PURTADS.Tables[0].Rows[0]["MV004"].ToString());
            return MV004;
        }

        public StringBuilder SETADDPURDTB(DATAPURTB DataPURTB)
        {
            StringBuilder ADDPURDTB = new StringBuilder();
            ADDPURDTB.AppendFormat(@"
                                    INSERT INTO [TK].dbo.PURTB
                                    (
                                    COMPANY,CREATOR,USR_GROUP,CREATE_DATE,MODIFIER,MODI_DATE,FLAG,CREATE_TIME,MODI_TIME,TRANS_TYPE,TRANS_NAME,sync_date,sync_time,sync_mark,sync_count,DataUser,DataGroup
                                    ,TB001,TB002,TB003,TB004,TB005,TB006,TB007,TB008,TB009,TB010
                                    ,TB011,TB012,TB013,TB014,TB015,TB016,TB017,TB018,TB019,TB020
                                    ,TB021,TB022,TB023,TB024,TB025,TB026,TB027,TB028,TB029,TB030
                                    ,TB031,TB032,TB033,TB034,TB035,TB036,TB037,TB038,TB039,TB040
                                    ,TB041,TB042,TB043,TB044,TB045,TB046,TB047,TB048,TB049,TB050
                                    ,TB051,TB052,TB053,TB054,TB055,TB056,TB057,TB058,TB059,TB060
                                    ,TB061,TB062,TB063,TB064,TB065,TB066,TB067,TB068,TB069,TB070
                                    ,TB071,TB072,TB073,TB074,TB075,TB076,TB077,TB078,TB079,TB080
                                    ,TB081,TB082,TB083,TB084,TB085,TB086,TB087,TB088,TB089,TB090
                                    ,TB091,TB092,TB093,TB094,TB095,TB096,TB097,TB098,TB099
                                    ,UDF01,UDF02,UDF03,UDF04,UDF05,UDF06,UDF07,UDF08,UDF09,UDF10
                                    )

                                   SELECT 
                                    '{1}'COMPANY,'{2}' CREATOR,'{3}' USR_GROUP,'{4}' CREATE_DATE,'{5}' MODIFIER,'{6}' MODI_DATE,'{7}'FLAG,'{8}' CREATE_TIME,'{9}' MODI_TIME,'{10}' TRANS_TYPE,'{11}' TRANS_NAME,'{12}' sync_date,'{13}' sync_time,'{14}' sync_mark,'{15}' sync_count,'{16}' DataUser,'{17}' DataGroup
                                    ,'{18}' TB001,'{19}' TB002,'{20}' TB003,MB001 TB004,MB002 TB005,MB003 TB006,MB004 TB007,MB017 TB008,'{21}' TB009,MB032 TB010
                                    ,'{22}' TB011,'{23}' TB012,MB067 TB013,'{21}' TB014,MB004 TB015,'NTD' TB016,CONVERT(DECIMAL(16,3),MB065/MB064) TB017,CONVERT(INT,MB065/MB064) TB018,'{22}' TB019,'N' TB020
                                    ,'N' TB021,'' TB022,'' TB023,'{24}' TB024,'N' TB025,CASE WHEN ISNULL(MA044,'')<>'' THEN MA044 ELSE '1' END TB026,'' TB027,'' TB028,'' TB029,'' TB030
                                    ,'' TB031,'N' TB032,'' TB033,0 TB034,0 TB035,'' TB036,'' TB037,'' TB038,'N' TB039,0 TB040
                                    ,0 TB041,'' TB042,'' TB043,'' TB044,'' TB045,'' TB046,'' TB047,'' TB048,0 TB049,'' TB050
                                    ,0 TB051,0 TB052,0 TB053,'' TB054,'' TB055,'' TB056,'' TB057,'1' TB058,'' TB059,'' TB060
                                    ,'' TB061,'' TB062,0 TB063,'N' TB064,'1' TB065,'' TB066,'2' TB067,0 TB068,0 TB069,'' TB070
                                    ,'' TB071,'' TB072,'' TB073,'' TB074,0 TB075,'' TB076,0 TB077,'' TB078,'' TB079,'' TB080
                                    ,0 TB081,0 TB082,0 TB083,0 TB084,0 TB085,'' TB086,'' TB087,'0' TB088,'1' TB089,0 TB090
                                    ,0 TB091,0 TB092,0 TB093,'' TB094,'' TB095,'' TB096,'' TB097,'' TB098,'' TB099
                                    ,'' UDF01,'' UDF02,'' UDF03,'' UDF04,'' UDF05,0 UDF06,0 UDF07,0 UDF08,0 UDF09,0 UDF10
                                    FROM [TK].dbo.INVMB
                                    LEFT JOIN [TK].dbo.PURMA ON MA001=MB032
                                    WHERE MB001 IN ('{0}')
                                    ", DataPURTB.TB004
                                    , DataPURTB.COMPANY, DataPURTB.CREATOR, DataPURTB.USR_GROUP, DataPURTB.CREATE_DATE, DataPURTB.MODIFIER, DataPURTB.MODI_DATE, DataPURTB.FLAG, DataPURTB.CREATE_TIME, DataPURTB.MODI_TIME, DataPURTB.TRANS_TYPE, DataPURTB.TRANS_NAME, DataPURTB.sync_date, DataPURTB.sync_time, DataPURTB.sync_mark, DataPURTB.sync_count, DataPURTB.DataUser, DataPURTB.DataGroup
                                    ,  DataPURTB.TB001, DataPURTB.TB002, DataPURTB.TB003, DataPURTB.TB009, DataPURTB.TB011, DataPURTB.TB012, DataPURTB.TB024);

            return ADDPURDTB;
        }

        public void INSERTINTOPURTAB(DATAPURTA DataPURTA, StringBuilder PURTBSB,string FormNumber)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                       INSERT INTO [TK].dbo.PURTA
                                        (
                                        COMPANY,CREATOR,USR_GROUP,CREATE_DATE,MODIFIER,MODI_DATE,FLAG,CREATE_TIME,MODI_TIME,TRANS_TYPE,TRANS_NAME,sync_date,sync_time,sync_mark,sync_count,DataUser,DataGroup,
                                        TA001,TA002,TA003,TA004,TA005,TA006,TA007,TA008,TA009,TA010,
                                        TA011,TA012,TA013,TA014,TA015,TA016,TA017,TA018,TA019,TA020,
                                        TA021,TA022,TA023,TA024,TA025,TA026,TA027,TA028,TA029,TA030,
                                        TA031,TA032,TA033,TA034,TA035,TA036,TA037,TA038,TA039,TA040,
                                        TA041,TA042,TA043,TA044,TA045,TA046,
                                        UDF01,UDF02,UDF03,UDF04,UDF05,UDF06,UDF07,UDF08,UDF09,UDF10
                                        )
                                        VALUES
                                        (
                                        @COMPANY,@CREATOR,@USR_GROUP,@CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG,@CREATE_TIME,@MODI_TIME,@TRANS_TYPE,@TRANS_NAME,@sync_date,@sync_time,@sync_mark,@sync_count,@DataUser,@DataGroup,
                                        @TA001,@TA002,@TA003,@TA004,@TA005,@TA006,@TA007,@TA008,@TA009,@TA010,
                                        @TA011,@TA012,@TA013,@TA014,@TA015,@TA016,@TA017,@TA018,@TA019,@TA020,
                                        @TA021,@TA022,@TA023,@TA024,@TA025,@TA026,@TA027,@TA028,@TA029,@TA030,
                                        @TA031,@TA032,@TA033,@TA034,@TA035,@TA036,@TA037,@TA038,@TA039,@TA040,
                                        @TA041,@TA042,@TA043,@TA044,@TA045,@TA046,
                                        @UDF01,@UDF02,@UDF03,@UDF04,@UDF05,@UDF06,@UDF07,@UDF08,@UDF09,@UDF10
                                        ) 
                                    ");
            queryString.AppendLine();

            queryString.Append(PURTBSB.ToString());

            queryString.AppendFormat(@"
                                    UPDATE [TK].dbo.PURTA
                                    SET TA011=(SELECT ISNULL(SUM(TB009),0) FROM [TK].dbo.PURTB WHERE TB001=@TB001 AND TB002=@TB002)
                                    WHERE TA001=@TA001 AND TA002=@TA002
                                    ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);

                    command.Parameters.Add("@COMPANY", SqlDbType.NVarChar).Value = DataPURTA.COMPANY;
                    command.Parameters.Add("@CREATOR", SqlDbType.NVarChar).Value = DataPURTA.CREATOR;
                    command.Parameters.Add("@USR_GROUP", SqlDbType.NVarChar).Value = DataPURTA.USR_GROUP;
                    command.Parameters.Add("@CREATE_DATE", SqlDbType.NVarChar).Value = DataPURTA.CREATE_DATE;
                    command.Parameters.Add("@MODIFIER", SqlDbType.NVarChar).Value = DataPURTA.MODIFIER;
                    command.Parameters.Add("@MODI_DATE", SqlDbType.NVarChar).Value = DataPURTA.MODI_DATE;
                    command.Parameters.Add("@FLAG", SqlDbType.NVarChar).Value = DataPURTA.FLAG;
                    command.Parameters.Add("@CREATE_TIME", SqlDbType.NVarChar).Value = DataPURTA.CREATE_TIME;
                    command.Parameters.Add("@MODI_TIME", SqlDbType.NVarChar).Value = DataPURTA.MODI_TIME;
                    command.Parameters.Add("@TRANS_TYPE", SqlDbType.NVarChar).Value = DataPURTA.TRANS_TYPE;
                    command.Parameters.Add("@TRANS_NAME", SqlDbType.NVarChar).Value = DataPURTA.TRANS_NAME;
                    command.Parameters.Add("@sync_date", SqlDbType.NVarChar).Value = DataPURTA.sync_date;
                    command.Parameters.Add("@sync_time", SqlDbType.NVarChar).Value = DataPURTA.sync_time;
                    command.Parameters.Add("@sync_mark", SqlDbType.NVarChar).Value = DataPURTA.sync_mark;
                    command.Parameters.Add("@sync_count", SqlDbType.NVarChar).Value = DataPURTA.sync_count;
                    command.Parameters.Add("@DataUser", SqlDbType.NVarChar).Value = DataPURTA.DataUser;
                    command.Parameters.Add("@DataGroup", SqlDbType.NVarChar).Value = DataPURTA.DataGroup;
                    command.Parameters.Add("@TA001", SqlDbType.NVarChar).Value = DataPURTA.TA001;
                    command.Parameters.Add("@TA002", SqlDbType.NVarChar).Value = DataPURTA.TA002;
                    command.Parameters.Add("@TA003", SqlDbType.NVarChar).Value = DataPURTA.TA003;
                    command.Parameters.Add("@TA004", SqlDbType.NVarChar).Value = DataPURTA.TA004;
                    command.Parameters.Add("@TA005", SqlDbType.NVarChar).Value = FormNumber;
                    command.Parameters.Add("@TA006", SqlDbType.NVarChar).Value = DataPURTA.TA006;
                    command.Parameters.Add("@TA007", SqlDbType.NVarChar).Value = DataPURTA.TA007;
                    command.Parameters.Add("@TA008", SqlDbType.NVarChar).Value = DataPURTA.TA008;
                    command.Parameters.Add("@TA009", SqlDbType.NVarChar).Value = DataPURTA.TA009;
                    command.Parameters.Add("@TA010", SqlDbType.NVarChar).Value = DataPURTA.TA010;
                    command.Parameters.Add("@TA011", SqlDbType.NVarChar).Value = DataPURTA.TA011;
                    command.Parameters.Add("@TA012", SqlDbType.NVarChar).Value = DataPURTA.TA012;
                    command.Parameters.Add("@TA013", SqlDbType.NVarChar).Value = DataPURTA.TA013;
                    command.Parameters.Add("@TA014", SqlDbType.NVarChar).Value = DataPURTA.TA014;
                    command.Parameters.Add("@TA015", SqlDbType.NVarChar).Value = DataPURTA.TA015;
                    command.Parameters.Add("@TA016", SqlDbType.NVarChar).Value = DataPURTA.TA016;
                    command.Parameters.Add("@TA017", SqlDbType.NVarChar).Value = DataPURTA.TA017;
                    command.Parameters.Add("@TA018", SqlDbType.NVarChar).Value = DataPURTA.TA018;
                    command.Parameters.Add("@TA019", SqlDbType.NVarChar).Value = DataPURTA.TA019;
                    command.Parameters.Add("@TA020", SqlDbType.NVarChar).Value = DataPURTA.TA020;
                    command.Parameters.Add("@TA021", SqlDbType.NVarChar).Value = DataPURTA.TA021;
                    command.Parameters.Add("@TA022", SqlDbType.NVarChar).Value = DataPURTA.TA022;
                    command.Parameters.Add("@TA023", SqlDbType.NVarChar).Value = DataPURTA.TA023;
                    command.Parameters.Add("@TA024", SqlDbType.NVarChar).Value = DataPURTA.TA024;
                    command.Parameters.Add("@TA025", SqlDbType.NVarChar).Value = DataPURTA.TA025;
                    command.Parameters.Add("@TA026", SqlDbType.NVarChar).Value = DataPURTA.TA026;
                    command.Parameters.Add("@TA027", SqlDbType.NVarChar).Value = DataPURTA.TA027;
                    command.Parameters.Add("@TA028", SqlDbType.NVarChar).Value = DataPURTA.TA028;
                    command.Parameters.Add("@TA029", SqlDbType.NVarChar).Value = DataPURTA.TA029;
                    command.Parameters.Add("@TA030", SqlDbType.NVarChar).Value = DataPURTA.TA030;
                    command.Parameters.Add("@TA031", SqlDbType.NVarChar).Value = DataPURTA.TA031;
                    command.Parameters.Add("@TA032", SqlDbType.NVarChar).Value = DataPURTA.TA032;
                    command.Parameters.Add("@TA033", SqlDbType.NVarChar).Value = DataPURTA.TA033;
                    command.Parameters.Add("@TA034", SqlDbType.NVarChar).Value = DataPURTA.TA034;
                    command.Parameters.Add("@TA035", SqlDbType.NVarChar).Value = DataPURTA.TA035;
                    command.Parameters.Add("@TA036", SqlDbType.NVarChar).Value = DataPURTA.TA036;
                    command.Parameters.Add("@TA037", SqlDbType.NVarChar).Value = DataPURTA.TA037;
                    command.Parameters.Add("@TA038", SqlDbType.NVarChar).Value = DataPURTA.TA038;
                    command.Parameters.Add("@TA039", SqlDbType.NVarChar).Value = DataPURTA.TA039;
                    command.Parameters.Add("@TA040", SqlDbType.NVarChar).Value = DataPURTA.TA040;
                    command.Parameters.Add("@TA041", SqlDbType.NVarChar).Value = DataPURTA.TA041;
                    command.Parameters.Add("@TA042", SqlDbType.NVarChar).Value = DataPURTA.TA042;
                    command.Parameters.Add("@TA043", SqlDbType.NVarChar).Value = DataPURTA.TA043;
                    command.Parameters.Add("@TA044", SqlDbType.NVarChar).Value = DataPURTA.TA044;
                    command.Parameters.Add("@TA045", SqlDbType.NVarChar).Value = DataPURTA.TA045;
                    command.Parameters.Add("@TA046", SqlDbType.NVarChar).Value = DataPURTA.TA046;
                    command.Parameters.Add("@UDF01", SqlDbType.NVarChar).Value = DataPURTA.UDF01;
                    command.Parameters.Add("@UDF02", SqlDbType.NVarChar).Value = DataPURTA.UDF02;
                    command.Parameters.Add("@UDF03", SqlDbType.NVarChar).Value = DataPURTA.UDF03;
                    command.Parameters.Add("@UDF04", SqlDbType.NVarChar).Value = DataPURTA.UDF04;
                    command.Parameters.Add("@UDF05", SqlDbType.NVarChar).Value = DataPURTA.UDF05;
                    command.Parameters.Add("@UDF06", SqlDbType.NVarChar).Value = DataPURTA.UDF06;
                    command.Parameters.Add("@UDF07", SqlDbType.NVarChar).Value = DataPURTA.UDF07;
                    command.Parameters.Add("@UDF08", SqlDbType.NVarChar).Value = DataPURTA.UDF08;
                    command.Parameters.Add("@UDF09", SqlDbType.NVarChar).Value = DataPURTA.UDF09;
                    command.Parameters.Add("@UDF10", SqlDbType.NVarChar).Value = DataPURTA.UDF10;


                    command.Parameters.Add("@TB001", SqlDbType.NVarChar).Value = DataPURTA.TA001;
                    command.Parameters.Add("@TB002", SqlDbType.NVarChar).Value = DataPURTA.TA002;

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

