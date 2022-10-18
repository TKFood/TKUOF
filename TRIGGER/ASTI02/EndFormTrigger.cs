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

namespace TKUOF.TRIGGER.ASTI02
{
    //訂單的核準

    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {

        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string MB001 = null;
            string MB016 = null;
            string MB042 = "AC01";
            string MB043 = null;
            string FORMID = null;
            string MODIFIER = null;
            UserUCO userUCO = new UserUCO();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            MB001 = applyTask.Task.CurrentDocument.Fields["MB001"].FieldValue.ToString().Trim();     
          
            FORMID = applyTask.FormNumber;
            //MODIFIER = applyTask.Task.Applicant.Account;

            //取得簽核人工號
            EBUser ebUser = userUCO.GetEBUser(Current.UserGUID);
            MODIFIER = ebUser.Account;

            //SEARCHASTMB
            DataTable DTSEARCHASTMB = SEARCHASTMB(MB001);
            MB016 = DTSEARCHASTMB.Rows[0]["MB016"].ToString();

            //取得MB042的單號
            MB043 = GETMAXNO(MB042,MB016);

            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(MB001) && !string.IsNullOrEmpty(MB042) && !string.IsNullOrEmpty(MB043) && !string.IsNullOrEmpty(MB016))
                {
                    UPDATEASTMBASTMC(MB001, FORMID, MODIFIER, MB042, MB043, MB016);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATEASTMBASTMC(string MB001, string FORMID, string MODIFIER, string MB042, string MB043,string MB016)
        {
            string TC027 = "Y";
            string TC048 = "N";
            string TD021 = "Y";
            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();



            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"                                   
                                        UPDATE [TK].dbo.ASTMB
                                        SET MB047=@MB047,MB048=@MB048
                                        ,MB042=@MB042,MB043=@MB043
                                        ,MB039=@MB039,MB050=@MB050
                                        ,FLAG=FLAG+1,COMPANY=@COMPANY,MODIFIER=@MODIFIER,MODI_DATE=@MODI_DATE,MODI_TIME=@MODI_TIME
                                        ,UDF02=@UDF02
                                        WHERE MB001=@MB001

                                        DELETE [TK].dbo.ASTTD
                                        WHERE TD001+TD002 IN (SELECT MB042+MB043 FROM [TK].dbo.ASTMB WHERE MB001=@MB001)

                                        DELETE [TK].dbo.ASTTC
                                        WHERE TC001+TC002 IN (SELECT MB042+MB043 FROM [TK].dbo.ASTMB WHERE MB001=@MB001)


                                        INSERT INTO [TK].dbo.ASTTC
                                        (
                                        TC001,TC002,TC003,TC004,TC005,TC006,TC007,TC008,TC009,TC010, 
                                        TC011,TC012,TC013,TC014,TC015,TC016,TC017,TC018,TC019,TC020,
                                        TC021,TC022,TC025,TC027,TC028,TC029,TC030,
                                        TC031,TC032,TC033,TC035,TC036,TC037,TC038,TC039, 
                                        TC068,TC069,
                                        TC073,TC074,TC075,TC076,TC077,TC078,TC079,TC080,
                                        TC086,TC088,TC092,TC093, 
                                        COMPANY,CREATOR,USR_GROUP,CREATE_DATE ,FLAG,
                                        CREATE_TIME,MODI_TIME,TRANS_TYPE,TRANS_NAME
                                        )
                                        SELECT 
                                        MB042 TC001,MB043 TC002,MB016 TC003,MB001 TC004,MB012 TC005,MB020 TC006,MB021 TC007,MB029 TC008,MB015 TC009,MB022 TC010, 
                                        MB052 TC011,MB053 TC012,MB032 TC013,'C0' TC014,'Y' TC015,0 TC016,'' TC017,'' TC018,MB018 TC019,MB080 TC020,
                                        0 TC021,0 TC022,'N' TC025,MB016 TC027,'AD' TC028,0 TC029,0 TC030,
                                        'N' TC031,'N' TC032,MB051 TC033,MB014 TC035,MB058 TC036,MB026 TC037,MB041 TC038,0 TC039, 
                                        MB076 TC068,'N' TC069,
                                        MB069 TC073,MB068 TC074,MB073 TC075,MB074 TC076, MB075 TC077,MB064 TC078,MB066 TC079,MB032 TC080,
                                        0 TC086,0 TC088,MB078 TC092,MB079 TC093, 
                                        'TK' COMPANY ,ASTMB.CREATOR CREATOR ,ASTMB.USR_GROUP USR_GROUP ,CONVERT(NVARCHAR,GETDATE(),112) CREATE_DATE ,0 FLAG,
                                        CONVERT(NVARCHAR,GETDATE(),108) CREATE_TIME,CONVERT(NVARCHAR,GETDATE(),108)  MODI_TIME, 'P004' TRANS_TYPE,'ASTI02' TRANS_NAME
                                        FROM [TK].dbo.ASTMB
                                        WHERE MB001=@MB001

                                        INSERT INTO [TK].dbo.ASTTD
                                        (
                                        TD001,TD002,TD003,TD004,TD005,TD006,TD007, TD008, TD009,TD010,TD016, 
                                        COMPANY,CREATOR,USR_GROUP,CREATE_DATE ,FLAG, CREATE_TIME
                                        )
                                        SELECT 
                                        MB042 TD001,MB043 TD002,MC002 TD003,MC003 TD004,MC004 TD005,MB031 TD006,0 TD007,'' TD008,'Y' TD009,MB001 TD010,MB066 TD016, 
                                        'TK' COMPANY ,ASTMB.CREATOR CREATOR ,ASTMB.USR_GROUP USR_GROUP ,CONVERT(NVARCHAR,GETDATE(),112) CREATE_DATE ,0 FLAG,
                                        CONVERT(NVARCHAR,GETDATE(),108) CREATE_TIME
                                        FROM [TK].dbo.ASTMB,[TK].dbo.ASTMC
                                        WHERE 1=1
                                        AND MB001=MC001
                                        AND MB001=@MB001

 
                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@MB001", SqlDbType.NVarChar).Value = MB001;
                    command.Parameters.Add("@MB047", SqlDbType.NVarChar).Value = MB016;
                    command.Parameters.Add("@MB048", SqlDbType.NVarChar).Value = MODIFIER;
                    command.Parameters.Add("@MB042", SqlDbType.NVarChar).Value = MB042;
                    command.Parameters.Add("@MB043", SqlDbType.NVarChar).Value = MB043;
                    command.Parameters.Add("@MB039", SqlDbType.NVarChar).Value = "Y";
                    command.Parameters.Add("@MB050", SqlDbType.NVarChar).Value = "N";
                    command.Parameters.Add("@COMPANY", SqlDbType.NVarChar).Value = "TK";
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


        public string GETMAXNO(string TC001,string TC003)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();
            Ede.Uof.Utility.Data.DatabaseHelper m_db = new Ede.Uof.Utility.Data.DatabaseHelper(connectionString);

            string cmdTxt = @" 
                            SELECT ISNULL(MAX(TC002),'00000000000') AS TA002
                            FROM  [TK].dbo.ASTTC
                            WHERE  TC001=@TC001 AND TC003=@TC003
                             ";

            m_db.AddParameter("@TC001", TC001);
            m_db.AddParameter("@TC003", TC003);


            DataTable dt = new DataTable();

            dt.Load(m_db.ExecuteReader(cmdTxt));

            if (dt.Rows.Count > 0)
            {
                string TA002 = SETTA002(dt.Rows[0]["TA002"].ToString(), TC003);
                return TA002;
            }
            else
            {
                return null;
            }
        }

        public string SETTA002(string TA002,string DATESTRING)
        {
            if (TA002.Equals("00000000000"))
            {
                return DATESTRING + "001";
            }

            else
            {
                int serno = Convert.ToInt16(TA002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return DATESTRING + temp.ToString();
            }

        }

        public DataTable SEARCHASTMB(string MB001)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();
            Ede.Uof.Utility.Data.DatabaseHelper m_db = new Ede.Uof.Utility.Data.DatabaseHelper(connectionString);

            string cmdTxt = @" 
                             SELECT MB016
                            FROM [TK].dbo.ASTMB
                            WHERE MB001=@MB001
                             ";

            m_db.AddParameter("@MB001", MB001);
        

            DataTable dt = new DataTable();

            dt.Load(m_db.ExecuteReader(cmdTxt));

            if (dt.Rows.Count > 0)
            {
                return dt;
            }
            else
            {
                return null;
            }
        }
    }
}
