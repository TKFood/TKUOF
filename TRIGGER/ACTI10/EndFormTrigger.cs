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

namespace TKUOF.TRIGGER.ACTI10
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
                    //會計傳票核準
                    UPDATE_ACTI10(TA001, TA002, FORMID, MODIFIER);
                    //過帳-分類帳檔
                    //ACTATB_TO_ACTML(TA001, TA002);
                    //會計科目各期額檔
                    //ACTATB_TO_ACTMB(TA001, TA002);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATE_ACTI10(string TA001, string TA002, string FORMID, string MODIFIER)
        {
            string TA014 = DateTime.Now.ToString("yyyyMMdd");
            string TA010 = "Y";
            string TA011 = "Y";
            string TA015 = MODIFIER;
            string TA016 = "N";
            string TA018 = "N";
            string TB016 = "Y";

            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();



            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"       
                                    
                                    UPDATE [test0923].dbo.ACTTA
                                    SET 
                                    TA014=@TA014 
                                    ,TA010=@TA010 
                                    ,TA011=@TA011
                                    ,TA015=@TA015 
                                    ,TA016=@TA016 
                                    ,TA018=@TA018 
                                    ,FLAG=FLAG+1
                                    ,MODIFIER=@MODIFIER 
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME
                                    WHERE TA001=@TA001 AND TA002=TA002

                                    UPDATE [test0923].dbo.ACTTB
                                    SET 
                                    TB016=@TB016 
                                    ,FLAG=FLAG+1
                                    ,MODIFIER=@MODIFIER 
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME
                                    WHERE TB001=@TB001 AND TB002=TB002

                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TA001", SqlDbType.NVarChar).Value = TA001;
                    command.Parameters.Add("@TA002", SqlDbType.NVarChar).Value = TA002;

                    command.Parameters.Add("@TB001", SqlDbType.NVarChar).Value = TA001;
                    command.Parameters.Add("@TB002", SqlDbType.NVarChar).Value = TA002;
                    command.Parameters.Add("@TA014", SqlDbType.NVarChar).Value = TA014;
                    command.Parameters.Add("@TA010", SqlDbType.NVarChar).Value = TA010;
                    command.Parameters.Add("@TA011", SqlDbType.NVarChar).Value = TA011;
                    command.Parameters.Add("@TA015", SqlDbType.NVarChar).Value = TA015;
                    command.Parameters.Add("@TA016", SqlDbType.NVarChar).Value = TA016;
                    command.Parameters.Add("@TA018", SqlDbType.NVarChar).Value = TA018;
                    command.Parameters.Add("@TB016", SqlDbType.NVarChar).Value = TB016;
           
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

        public void ACTATB_TO_ACTML(string TA001,string TA002)
        {
            string TA014 = DateTime.Now.ToString("yyyyMMdd");
          
            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"       
                                    INSERT INTO [test0923].dbo.ACTML
                                    (ML001, ML002, ML003, ML004, ML005, ML006,  ML007, ML008, ML009, ML010, ML011,ML012,ML013,ML014 ,ML022,ML023,ML024,ML025,ML026,ML027,ML020,ML034 ,COMPANY,CREATOR,USR_GROUP,CREATE_DATE,FLAG,CREATE_TIME,MODI_TIME,TRANS_TYPE,TRANS_NAME)
                                    
                                    SELECT 
                                    (CASE WHEN MA008='3' THEN TB005 ELSE (SELECT TOP 1 MA001 FROM [test0923].dbo.ACTMA MA1 WHERE MA1.MA008='1' AND MA1.MA002=SUBSTRING(ACTMA.MA002,1,LEN(ACTMA.MA002)-2)) END ) ML001
                                    ,TA003 ML002,TB001 ML003,TB002 ML004,TB003 ML005,TB005 ML006,TB004  ML007,TB007 ML008,TB010 ML009,TB006 ML010
                                    ,TB011 ML011,TB013 ML012,TB014 ML013,TB015 ML014 ,TB035 ML022,TB036 ML023,TB037 ML024,TB038 ML025
                                    ,TB039 ML026,TB040 ML027,TB033 ML020,TA041 ML034 
                                    ,'TK' COMPANY,ACTTA.CREATOR CREATOR,ACTTA.USR_GROUP USR_GROUP,CONVERT(NVARCHAR,GETDATE(),112) CREATE_DATE,2 FLAG,CONVERT(NVARCHAR,GETDATE(),108) CREATE_TIME,CONVERT(NVARCHAR,GETDATE(),108) MODI_TIME,'P003' TRANS_TYPE,'Actb03' TRANS_NAME
                                    FROM [test0923].dbo.ACTTB,[test0923].dbo.ACTTA,[test0923].dbo.ACTMA
                                    WHERE TA001=TB001 AND TA002=TB002
                                    AND TB005=MA001
                                    AND TB001=@TB001 AND TB002=@TB002

                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TB001", SqlDbType.NVarChar).Value = TA001;
                    command.Parameters.Add("@TB002", SqlDbType.NVarChar).Value = TA002;

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

        /// <summary>
        /// 找出會計傳票中所有的科目跟統制科目的借、袋方金額
        //  再判斷該年、該月是否有資料存在ACTMB
        //  有就UPDATE
        //  沒有就INSERT
        /// </summary>
        /// <param name="TA001"></param>
        /// <param name="TA002"></param>
        public void ACTATB_TO_ACTMB(string TA001, string TA002)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();
            StringBuilder queryString = new StringBuilder();

            DataTable DT = FIND_ACTATB_TO_ACTMB(TA001, TA002);

            if(DT!=null && DT.Rows.Count>=1)
            {               
                queryString.AppendFormat(@"  
                                        ");

                foreach(DataRow DR in DT.Rows)
                {
                    queryString.AppendFormat(@"  
                                            IF EXISTS (SELECT MB001 FROM [test0923].dbo.[ACTMB] WHERE MB001=@MB001 AND MB002=@MB002 AND MB003=@MB003 AND MB008=@MB008)
                                            BEGIN
	                                            UPDATE  [test0923].dbo.[ACTMB]
	                                            SET MB004=MB004+MB004NEW,MB005=MB005+MB005NEW,MB006=MB006+MB006NEW,MB007=MB007+MB007NEW,MB009=MB009+MB009NEW,MB010=MB010+MB010NEW
	                                            WHERE MB001=@MB001 AND MB002=@MB002 AND MB003=@MB003
                                            END

                                            IF NOT EXISTS (SELECT MB001 FROM [test0923].dbo.[ACTMB] WHERE MB001=@MB001 AND MB002=@MB002 AND MB003=@MB003)
                                            BEGIN
	                                            INSERT INTO  [test0923].dbo.[ACTMB]
	                                            (MB001,MB002,MB003,MB004,MB005,MB006,MB007,MB008,MB009,MB010,COMPANY,CREATOR,USR_GROUP,CREATE_DATE,MODIFIER,MODI_DATE,FLAG,CREATE_TIME,MODI_TIME,TRANS_TYPE,TRANS_NAME)
	                                            VALUE
	                                            (@MB001,@MB002,@MB003,@MB004,@MB005,@MB006,@MB007,@MB008,@MB009,@MB010,@COMPANY,@CREATOR,@USR_GROUP,@CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG,@CREATE_TIME,@MODI_TIME,@TRANS_TYPE,@TRANS_NAME)
                                            END

                                             ");

                    try
                    {
                        using (SqlConnection connection = new SqlConnection(connectionString))
                        {

                            SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                            command.Parameters.Add("@MB001", SqlDbType.NVarChar).Value = DR["TB005"].ToString();
                            command.Parameters.Add("@MB002", SqlDbType.NVarChar).Value = DR["YEARS"].ToString();
                            command.Parameters.Add("@MB003", SqlDbType.NVarChar).Value = DR["MONTHS"].ToString();
                            command.Parameters.Add("@MB008", SqlDbType.NVarChar).Value = DR["TB013"].ToString();
                            command.Parameters.Add("@MB004NEW", SqlDbType.NVarChar).Value = DR["MB004NEW"].ToString();
                            command.Parameters.Add("@MB005NEW", SqlDbType.NVarChar).Value = DR["MB005NEW"].ToString();
                            command.Parameters.Add("@MB006NEW", SqlDbType.NVarChar).Value = DR["MB006NEW"].ToString();
                            command.Parameters.Add("@MB007NEW", SqlDbType.NVarChar).Value = DR["MB007NEW"].ToString();
                            command.Parameters.Add("@MB009NEW", SqlDbType.NVarChar).Value = DR["MB009NEW"].ToString();
                            command.Parameters.Add("@MB010NEW", SqlDbType.NVarChar).Value = DR["MB010NEW"].ToString();

                 
                            command.Parameters.Add("@MB004", SqlDbType.NVarChar).Value = DR["MB004"].ToString();
                            command.Parameters.Add("@MB005", SqlDbType.NVarChar).Value = DR["MB005"].ToString();
                            command.Parameters.Add("@MB006", SqlDbType.NVarChar).Value = DR["MB006"].ToString();
                            command.Parameters.Add("@MB007", SqlDbType.NVarChar).Value = DR["MB007"].ToString();
                            command.Parameters.Add("@MB009", SqlDbType.NVarChar).Value = DR["MB009"].ToString();
                            command.Parameters.Add("@MB010", SqlDbType.NVarChar).Value = DR["MB010"].ToString();
                            command.Parameters.Add("@COMPANY", SqlDbType.NVarChar).Value = DR["COMPANY"].ToString();
                            command.Parameters.Add("@CREATOR", SqlDbType.NVarChar).Value = DR["CREATOR"].ToString();
                            command.Parameters.Add("@USR_GROUP", SqlDbType.NVarChar).Value = DR["USR_GROUP"].ToString();
                            command.Parameters.Add("@CREATE_DATE", SqlDbType.NVarChar).Value = DR["CREATE_DATE"].ToString();
                            command.Parameters.Add("@MODIFIER", SqlDbType.NVarChar).Value = DR["MODIFIER"].ToString();
                            command.Parameters.Add("@MODI_DATE", SqlDbType.NVarChar).Value = DR["MODI_DATE"].ToString();
                            command.Parameters.Add("@FLAG", SqlDbType.NVarChar).Value = DR["FLAG"].ToString();
                            command.Parameters.Add("@CREATE_TIME", SqlDbType.NVarChar).Value = DR["CREATE_TIME"].ToString();
                            command.Parameters.Add("@MODI_TIME", SqlDbType.NVarChar).Value = DR["MODI_TIME"].ToString();
                            command.Parameters.Add("@TRANS_TYPE", SqlDbType.NVarChar).Value = DR["TRANS_TYPE"].ToString();
                            command.Parameters.Add("@TRANS_NAME", SqlDbType.NVarChar).Value = DR["TRANS_NAME"].ToString();

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

        public DataTable FIND_ACTATB_TO_ACTMB(string TA001, string TA002)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();
            Ede.Uof.Utility.Data.DatabaseHelper m_db = new Ede.Uof.Utility.Data.DatabaseHelper(connectionString);

            string cmdTxt = @" 
                            SELECT *
                            FROM 
                            (
                            SELECT TB004,TB005,TB006,TB007,TB013,TB015,SUBSTRING(TA003,1,4) AS 'YEARS',SUBSTRING(TA003,5,2) AS 'MONTHS'
                            ,(CASE WHEN TB004='1' THEN 1 ELSE 0 END) MB006NEW
                            ,(CASE WHEN TB004='-1' THEN 1 ELSE 0 END) MB007NEW
                            ,(CASE WHEN TB004='1' THEN TB007 ELSE 0 END) MB004NEW
                            ,(CASE WHEN TB004='-1' THEN TB007 ELSE 0 END) MB005NEW
                            ,(CASE WHEN TB004='1' THEN TB015 ELSE 0 END) MB009NEW
                            ,(CASE WHEN TB004='-1' THEN TB015 ELSE 0 END) MB010NEW
                            ,ACTTA.COMPANY,ACTTA.CREATOR,ACTTA.USR_GROUP,ACTTA.CREATE_DATE,ACTTA.MODIFIER,ACTTA.MODI_DATE,0 AS FLAS,ACTTA.CREATE_TIME,ACTTA.MODI_TIME,'P003' TRANS_TYPE,'Actb03'TRANS_NAME
                            FROM [test0923].dbo.ACTTA,[test0923].dbo.ACTTB
                            WHERE TA001=TB001 AND TA002=TB002
                            AND TA001=@TA001 AND TA002=@TA002'

                            UNION ALL
                            SELECT TB004,MA2.MA001,TB006,TB007,TB013,TB015,SUBSTRING(TA003,1,4) AS 'YEARS',SUBSTRING(TA003,5,2) AS 'MONTHS'
                            ,(CASE WHEN TB004='1' THEN 1 ELSE 0 END) MB006NEW
                            ,(CASE WHEN TB004='-1' THEN 1 ELSE 0 END) MB007NEW
                            ,(CASE WHEN TB004='1' THEN TB007 ELSE 0 END) MB004NEW
                            ,(CASE WHEN TB004='-1' THEN TB007 ELSE 0 END) MB005NEW
                            ,(CASE WHEN TB004='1' THEN TB015 ELSE 0 END) MB009NEW
                            ,(CASE WHEN TB004='-1' THEN TB015 ELSE 0 END) MB010NEW
                            ,ACTTA.COMPANY,ACTTA.CREATOR,ACTTA.USR_GROUP,ACTTA.CREATE_DATE,ACTTA.MODIFIER,ACTTA.MODI_DATE,0 AS FLAS,ACTTA.CREATE_TIME,ACTTA.MODI_TIME,'P003' TRANS_TYPE,'Actb03'TRANS_NAME
                            FROM [test0923].dbo.ACTTA,[test0923].dbo.ACTTB,[test0923].dbo.ACTMA MA1,[test0923].dbo.ACTMA MA2
                            WHERE TA001=TB001 AND TA002=TB002
                            AND TA001=@TA001 AND TA002=@TA002'
                            AND TB005=MA1.MA001
                            AND MA1.MA008<>'3'
                            AND MA2.MA002=SUBSTRING(MA1.MA002,1,LEN(MA1.MA002)-2)
                            ) AS TEMP

                             ";

            m_db.AddParameter("@TA001", TA001);
            m_db.AddParameter("@TA002", TA002);


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
