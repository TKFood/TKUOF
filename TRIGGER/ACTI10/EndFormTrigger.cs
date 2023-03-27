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
                    ACTATB_TO_ACTML(TA001, TA002);
                    //會計科目各期額檔
                    ACTATB_TO_ACTMB(TA001, TA002);
                    //ACTMD 部門科目各期金額檔
                    ACTATB_TO_ACTMD(TA001, TA002);
                    //ACTMM 立沖帳目金額檔
                    ACTATB_TO_ACTMM(TA001, TA002);
                    //ACTMN 立沖帳目來源檔
                    ACTATB_TO_ACTMN(TA001, TA002);

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

                foreach (DataRow DR in DT.Rows)
                {
                    queryString.AppendFormat(@"  
                          
                                                IF NOT EXISTS (SELECT MB001 FROM [test0923].dbo.[ACTMB] WHERE MB001='{0}' AND MB002='{1}' AND MB003='{2}' AND MB008='{3}')
                                                BEGIN
                                                 INSERT INTO  [test0923].dbo.[ACTMB]
                                                 (MB001,MB002,MB003,MB004,MB005,MB006,MB007,MB008,MB009,MB010,COMPANY,CREATOR,USR_GROUP,CREATE_DATE,MODIFIER,MODI_DATE,FLAG,CREATE_TIME,MODI_TIME,TRANS_TYPE,TRANS_NAME)
                                                 VALUES
                                                 ('{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}','{24}')
                                                END                                       

                                                 ", DR["MB001"].ToString().Trim(), DR["MB002"].ToString().Trim(), DR["MB003"].ToString().Trim(), DR["MB008NEW"].ToString().Trim()
                                                  , DR["MB001"].ToString().Trim(), DR["MB002"].ToString().Trim(), DR["MB003"].ToString().Trim(), DR["MB004NEW"].ToString().Trim(), DR["MB005NEW"].ToString().Trim(), DR["MB006NEW"].ToString().Trim(), DR["MB007NEW"].ToString(), DR["MB008NEW"].ToString().Trim(), DR["MB009NEW"].ToString().Trim(), DR["MB010NEW"].ToString().Trim()
                                                  , DR["COMPANY"].ToString().Trim(), DR["CREATOR"].ToString().Trim(), DR["USR_GROUP"].ToString().Trim(), DR["CREATE_DATE"].ToString().Trim(), DR["MODIFIER"].ToString().Trim(), DR["MODI_DATE"].ToString().Trim(), '0', DR["CREATE_TIME"].ToString().Trim(), DR["MODI_TIME"].ToString().Trim(), DR["TRANS_TYPE"].ToString().Trim(), DR["TRANS_NAME"].ToString().Trim()
                                                 );
                    queryString.AppendFormat(@"  
                          
                                                IF  EXISTS (SELECT MB001 FROM [test0923].dbo.[ACTMB] WHERE MB001='{0}' AND MB002='{1}' AND MB003='{2}'  AND MB008='{3}')
                                                BEGIN
                                                 UPDATE  [test0923].dbo.[ACTMB]
                                                 SET MB004=MB004+{4},MB005=MB005+{5},MB006=MB006+{6},MB007=MB007+{7},MB009=MB009+{8},MB010=MB010+{9}
                                                 WHERE MB001='{0}' AND MB002='{1}' AND MB003='{2}'  AND MB008='{3}'
                                                END

                                                 ", DR["MB001"].ToString().Trim(), DR["MB002"].ToString().Trim(), DR["MB003"].ToString().Trim(), DR["MB008NEW"].ToString().Trim()
                                                ,  DR["MB004NEW"].ToString().Trim(), DR["MB005NEW"].ToString().Trim(), DR["MB006NEW"].ToString().Trim(), DR["MB007NEW"].ToString().Trim(), DR["MB009NEW"].ToString().Trim(), DR["MB010NEW"].ToString().Trim()

                                               );

                }

                try
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {

                        SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                        //command.Parameters.Add("@MB001", SqlDbType.NVarChar).Value = DR["TB005"].ToString();                           
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

        public DataTable FIND_ACTATB_TO_ACTMB(string TA001, string TA002)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();
            Ede.Uof.Utility.Data.DatabaseHelper m_db = new Ede.Uof.Utility.Data.DatabaseHelper(connectionString);

            string cmdTxt = @" 
                            SELECT *
                            FROM 
                            (
                            SELECT TB004,TB005 AS MB001 ,TB006,TB007,TB013  AS 'MB008NEW',TB015,SUBSTRING(TA003,1,4) AS 'MB002',SUBSTRING(TA003,5,2) AS 'MB003'
                            ,(CASE WHEN TB004='1' THEN 1 ELSE 0 END) MB006NEW
                            ,(CASE WHEN TB004='-1' THEN 1 ELSE 0 END) MB007NEW
                            ,(CASE WHEN TB004='1' THEN TB007 ELSE 0 END) MB004NEW
                            ,(CASE WHEN TB004='-1' THEN TB007 ELSE 0 END) MB005NEW
                            ,(CASE WHEN TB004='1' THEN TB015 ELSE 0 END) MB009NEW
                            ,(CASE WHEN TB004='-1' THEN TB015 ELSE 0 END) MB010NEW
                            ,ACTTA.COMPANY,ACTTA.CREATOR,ACTTA.USR_GROUP,ACTTA.CREATE_DATE,ACTTA.MODIFIER,ACTTA.MODI_DATE,0 AS FLAS,ACTTA.CREATE_TIME,ACTTA.MODI_TIME,'P003' TRANS_TYPE,'Actb03'TRANS_NAME
                            FROM [test0923].dbo.ACTTA,[test0923].dbo.ACTTB
                            WHERE TA001=TB001 AND TA002=TB002
                            AND TA001=@TA001 AND TA002=@TA002

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
                            AND TA001=@TA001 AND TA002=@TA002
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



        /// <summary>
        /// 找出會計傳票中所有的科目跟統制科目的借、袋方金額
        //  再判斷該年、ACTATB_TO_ACTMD
        //  有就UPDATE
        //  沒有就INSERT
        /// </summary>
        /// <param name="TA001"></param>
        /// <param name="TA002"></param>
        public void ACTATB_TO_ACTMD(string TA001, string TA002)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();
            StringBuilder queryString = new StringBuilder();

            DataTable DT = FIND_ACTATB_TO_ACTMD(TA001, TA002);

            if (DT != null && DT.Rows.Count >= 1)
            {
                queryString.AppendFormat(@"  
                                        ");

                foreach (DataRow DR in DT.Rows)
                {
                    queryString.AppendFormat(@"  
                                            IF NOT EXISTS (SELECT MD001 FROM [test0923].dbo.[ACTMD] WHERE MD001='{0}' AND MD002='{1}' AND MD003='{2}' AND MD004='{3}' AND MD007='{4}')
                                            BEGIN
	                                            INSERT INTO  [test0923].dbo.[ACTMD]
	                                            (MD001,MD002,MD003,MD004,MD005,MD006,MD007,MD008,MD009,COMPANY,CREATOR,USR_GROUP,CREATE_DATE,MODIFIER,MODI_DATE,FLAG,CREATE_TIME,MODI_TIME,TRANS_TYPE,TRANS_NAME)
	                                            VALUES
	                                            ('{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}','{24}')
                                            END
                                             ", DR["MD001"].ToString().Trim(), DR["MD002"].ToString().Trim(), DR["MD003"].ToString().Trim(), DR["MD004"].ToString().Trim(), DR["MD007"].ToString().Trim()
                                             , DR["MD001"].ToString().Trim(), DR["MD002"].ToString().Trim(), DR["MD003"].ToString().Trim(), DR["MD004"].ToString().Trim(), DR["MD005NEW"].ToString().Trim(), DR["MD006NEW"].ToString().Trim(), DR["MD007"].ToString().Trim(), DR["MD008NEW"].ToString().Trim(), DR["MD009NEW"].ToString().Trim()
                                             , DR["COMPANY"].ToString().Trim(), DR["CREATOR"].ToString().Trim(), DR["USR_GROUP"].ToString().Trim(), DR["CREATE_DATE"].ToString().Trim(), DR["MODIFIER"].ToString().Trim(), DR["MODI_DATE"].ToString().Trim(), '0', DR["CREATE_TIME"].ToString().Trim(), DR["MODI_TIME"].ToString().Trim(), DR["TRANS_TYPE"].ToString().Trim(), DR["TRANS_NAME"].ToString().Trim()
                                             );

                    queryString.AppendFormat(@"  
                                            IF  EXISTS (SELECT MD001 FROM [test0923].dbo.[ACTMD] WHERE MD001='{0}' AND MD002='{1}' AND MD003='{2}' AND MD004='{3}' AND MD007='{4}')
                                            BEGIN
	                                           UPDATE [test0923].dbo.[ACTMD]
                                                SET MD005=MD005+{5},MD006=MD006+{6},MD008=MD008+{7},MD009=MD009+{8}
                                                WHERE  MD001='{0}' AND MD002='{1}' AND MD003='{2}' AND MD004='{3}' AND MD007='{4}'
                                            END
                                             ", DR["MD001"].ToString().Trim(), DR["MD002"].ToString(), DR["MD003"].ToString().Trim(), DR["MD004"].ToString().Trim(), DR["MD007"].ToString().Trim()
                                            , DR["MD005NEW"].ToString().Trim(), DR["MD006NEW"].ToString().Trim(), DR["MD008NEW"].ToString().Trim(), DR["MD009NEW"].ToString().Trim()

                                            );

                }

                try
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {

                        SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                        //command.Parameters.Add("@MB001", SqlDbType.NVarChar).Value = DR["TB005"].ToString();                           
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

        public DataTable FIND_ACTATB_TO_ACTMD(string TA001, string TA002)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();
            Ede.Uof.Utility.Data.DatabaseHelper m_db = new Ede.Uof.Utility.Data.DatabaseHelper(connectionString);

            string cmdTxt = @" 
                            SELECT *
                            FROM 
                            (
                            SELECT TB004,TB005 AS MD001 ,TB006 MD002,TB007,TB013  AS 'MD007',TB015,SUBSTRING(TA003,1,4) AS 'MD003',SUBSTRING(TA003,5,2) AS 'MD004'
                            ,(CASE WHEN TB004='1' THEN 1 ELSE 0 END) MB006NEW
                            ,(CASE WHEN TB004='-1' THEN 1 ELSE 0 END) MB007NEW
                            ,(CASE WHEN TB004='1' THEN TB007 ELSE 0 END) MD005NEW
                            ,(CASE WHEN TB004='-1' THEN TB007 ELSE 0 END) MD006NEW
                            ,(CASE WHEN TB004='1' THEN TB015 ELSE 0 END) MD008NEW
                            ,(CASE WHEN TB004='-1' THEN TB015 ELSE 0 END) MD009NEW
                            ,ACTTA.COMPANY,ACTTA.CREATOR,ACTTA.USR_GROUP,ACTTA.CREATE_DATE,ACTTA.MODIFIER,ACTTA.MODI_DATE,0 AS FLAS,ACTTA.CREATE_TIME,ACTTA.MODI_TIME,'P003' TRANS_TYPE,'Actb03'TRANS_NAME
                            FROM [test0923].dbo.ACTTA,[test0923].dbo.ACTTB
                            WHERE TA001=TB001 AND TA002=TB002
                            AND TA001=@TA001 AND TA002=@TA002

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
                            AND TA001=@TA001 AND TA002=@TA002
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



        /// <summary>
        /// 找出會計傳票中ACTMM 立沖帳目金額檔
        //  再判斷該年、月的ACTATB_TO_ACTMM
        //  有就UPDATE
        //  沒有就INSERT
        /// </summary>
        /// <param name="TA001"></param>
        /// <param name="TA002"></param>
        public void ACTATB_TO_ACTMM(string TA001, string TA002)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();
            StringBuilder queryString = new StringBuilder();

            DataTable DT = FIND_ACTATB_TO_ACTMM(TA001, TA002);

            if (DT != null && DT.Rows.Count >= 1)
            {
                queryString.AppendFormat(@"  
                                        ");

                foreach (DataRow DR in DT.Rows)
                {
                    queryString.AppendFormat(@"  
                                                IF NOT EXISTS (SELECT MM001 FROM [test0923].dbo.[ACTMM] WHERE MM001='{0}' AND MM002='{1}' AND MM003='{2}' AND MM004='{3}' AND MM005='{4}' AND MM016='{5}')
                                                BEGIN
	                                                INSERT INTO  [test0923].dbo.[ACTMM]
	                                                (MM001,MM002,MM003,MM004,MM005,MM006,MM007,MM008,MM009,MM016,COMPANY,CREATOR,USR_GROUP,CREATE_DATE,MODIFIER,MODI_DATE,FLAG,CREATE_TIME,MODI_TIME,TRANS_TYPE,TRANS_NAME)
	                                                VALUES
	                                                ('{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}','{24}','{25}','{26}')
                                                END
                                       
                                                ", DR["MM001"].ToString().Trim(), DR["MM002"].ToString().Trim(), DR["MM003"].ToString().Trim(), DR["MM004"].ToString().Trim(), DR["MM005"].ToString().Trim(), DR["MM016"].ToString().Trim()
                                                , DR["MM001"].ToString().Trim(), DR["MM002"].ToString().Trim(), DR["MM003"].ToString().Trim(), DR["MM004"].ToString().Trim(), DR["MM005"].ToString().Trim(), DR["MM006NEW"].ToString().Trim(), DR["MM007NEW"].ToString().Trim(), DR["MM008NEW"].ToString().Trim(), DR["MM009NEW"].ToString().Trim(), DR["MM016"].ToString().Trim()
                                                , DR["COMPANY"].ToString().Trim(), DR["CREATOR"].ToString().Trim(), DR["USR_GROUP"].ToString().Trim(), DR["CREATE_DATE"].ToString().Trim(), DR["MODIFIER"].ToString().Trim(), DR["MODI_DATE"].ToString().Trim(), '0', DR["CREATE_TIME"].ToString().Trim(), DR["MODI_TIME"].ToString().Trim(), DR["TRANS_TYPE"].ToString().Trim(), DR["TRANS_NAME"].ToString().Trim()
                                                );

                    queryString.AppendFormat(@"  
                                                IF  EXISTS (SELECT MM001 FROM [test0923].dbo.[ACTMM] WHERE MM001='{0}' AND MM002='{1}' AND MM003='{2}' AND MM004='{3}' AND MM005='{4}' AND MM016='{5}')
                                                BEGIN
	                                                UPDATE [test0923].dbo.[ACTMM]
	                                                SET MM006=MM006+{6},MM007=MM007+{7},MM008=MM008+{8},MM009=MM009+{9}
                                                    WHERE  MM001='{0}' AND MM002='{1}' AND MM003='{2}' AND MM004='{3}' AND MM005='{4}' AND MM016='{5}'
                                                END
                                       
                                                ", DR["MM001"].ToString().Trim(), DR["MM002"].ToString().Trim(), DR["MM003"].ToString().Trim(), DR["MM004"].ToString().Trim(), DR["MM005"].ToString().Trim(), DR["MM016"].ToString().Trim()
                                               ,  DR["MM006NEW"].ToString().Trim(), DR["MM007NEW"].ToString().Trim(), DR["MM008NEW"].ToString().Trim(), DR["MM009NEW"].ToString().Trim()
                                               
                                               );



                }

                try
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {

                        SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                        //command.Parameters.Add("@MB001", SqlDbType.NVarChar).Value = DR["TB005"].ToString();                           
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

        public DataTable FIND_ACTATB_TO_ACTMM(string TA001, string TA002)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();
            Ede.Uof.Utility.Data.DatabaseHelper m_db = new Ede.Uof.Utility.Data.DatabaseHelper(connectionString);

            string cmdTxt = @"                             
                                SELECT *
                                FROM
                                (
                                SELECT TB004,TB005 MM001,'1' MM002,TB008 MM003,TB006,TB007,TB013,TB015,SUBSTRING(TA003,1,4) AS 'MM004',SUBSTRING(TA003,5,2) AS 'MM005',TA041 AS MM016
                                ,(CASE WHEN TB004='1' THEN 1 ELSE 0 END) MM008NEW
                                ,(CASE WHEN TB004='-1' THEN 1 ELSE 0 END) MM009NEW
                                ,(CASE WHEN TB004='1' THEN TB007 ELSE 0 END) MM006NEW
                                ,(CASE WHEN TB004='-1' THEN TB007 ELSE 0 END) MM007NEW
                                ,ACTTA.COMPANY,ACTTA.CREATOR,ACTTA.USR_GROUP,ACTTA.CREATE_DATE,ACTTA.MODIFIER,ACTTA.MODI_DATE,0 AS FLAS,ACTTA.CREATE_TIME,ACTTA.MODI_TIME,'P003' TRANS_TYPE,'Actb03'TRANS_NAME

                                FROM [test0923].dbo.ACTTA,[test0923].dbo.ACTTB
                                WHERE TA001=TB001 AND TA002=TB002
                                AND (ISNULL(TB008,'')<>'')
                                AND TA001=@TA001 AND TA002=@TA002
                                UNION ALL
                                SELECT TB004,TB005,'2' MM002,TB009 MM003,TB006,TB007,TB013,TB015,SUBSTRING(TA003,1,4) AS 'YEARS',SUBSTRING(TA003,5,2) AS 'MONTHS',TA041 AS MM016
                                ,(CASE WHEN TB004='1' THEN 1 ELSE 0 END) MM008NEW
                                ,(CASE WHEN TB004='-1' THEN 1 ELSE 0 END) MM009NEW
                                ,(CASE WHEN TB004='1' THEN TB007 ELSE 0 END) MB004NEW
                                ,(CASE WHEN TB004='-1' THEN TB007 ELSE 0 END) MB005NEW
                                ,ACTTA.COMPANY,ACTTA.CREATOR,ACTTA.USR_GROUP,ACTTA.CREATE_DATE,ACTTA.MODIFIER,ACTTA.MODI_DATE,0 AS FLAS,ACTTA.CREATE_TIME,ACTTA.MODI_TIME,'P003' TRANS_TYPE,'Actb03'TRANS_NAME

                                FROM [test0923].dbo.ACTTA,[test0923].dbo.ACTTB
                                WHERE TA001=TB001 AND TA002=TB002
                                AND (ISNULL(TB009,'')<>'')
                                AND TA001=@TA001 AND TA002=@TA002
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

        /// <summary>
        /// 找出會計傳票中ACTMN 立沖帳目來源檔
        //  再判斷該年、月的ACTATB_TO_ACTMN 立沖帳目來源檔
        //  有就UPDATE
        //  沒有就INSERT
        /// </summary>
        /// <param name="TA001"></param>
        /// <param name="TA002"></param>
        public void ACTATB_TO_ACTMN(string TA001, string TA002)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();
            StringBuilder queryString = new StringBuilder();

            queryString.AppendFormat(@"  
                                    INSERT INTO [test0923].dbo.ACTMN
                                    (MN001,MN002,MN003,MN004,MN005,MN006,MN007,MN008,MN009,MN010,MN011,MN012,MN013,MN014,MN015,MN018,MN019,MN020,MN027
                                    ,COMPANY,CREATOR,USR_GROUP,CREATE_DATE,MODIFIER,MODI_DATE,FLAG,CREATE_TIME,MODI_TIME,TRANS_TYPE,TRANS_NAME)

                                    SELECT MN001,MN002,MN003,MN004,MN005,MN006,MN007,MN008,MN009,MN010,MN011,MN012,MN013,MN014,MN015,MN018,MN019,MN020,MN027
                                            ,COMPANY,CREATOR,USR_GROUP,CREATE_DATE,MODIFIER,MODI_DATE,FLAG,CREATE_TIME,MODI_TIME,TRANS_TYPE,TRANS_NAME
                                    FROM
                                    (
                                    SELECT TB005 MN001,'1' AS MN002,TB008 MN003,TA003 MN004,TB001 MN005,TB002 MN006,TB003 MN007,TB004 MN008,TB007 MN009,TB010 MN010,TB006 MN011,TB013 MN012,TB014 MN013,TB015 MN014,TB017 MN015,0 MN018,0 MN019,'N' MN020,TA041 MN027
                                    ,ACTTA.COMPANY,ACTTA.CREATOR,ACTTA.USR_GROUP,ACTTA.CREATE_DATE,ACTTA.MODIFIER,ACTTA.MODI_DATE,0 AS FLAG,ACTTA.CREATE_TIME,ACTTA.MODI_TIME,'P003' TRANS_TYPE,'Actb03'TRANS_NAME

                                    FROM [test0923].dbo.ACTTA,[test0923].dbo.ACTTB
                                    WHERE TA001=TB001 AND TA002=TB002
                                    AND (ISNULL(TB008,'')<>'')
                                    AND ISNULL(TB021,'')=''
                                    AND TA001=@TA001 AND TA002=@TA002
                                    UNION ALL
                                    SELECT TB005 MN001,'1' AS MN002,TB008 MN003,TA003 MN004,TB001 MN005,TB002 MN006,TB003 MN007,TB004 MN008,TB007 MN009,TB010 MN010,TB006 MN011,TB013 MN012,TB014 MN013,TB015 MN014,TB017 MN015,0 MN018,0 MN019,'N' MN020,TA041 MN027
                                    ,ACTTA.COMPANY,ACTTA.CREATOR,ACTTA.USR_GROUP,ACTTA.CREATE_DATE,ACTTA.MODIFIER,ACTTA.MODI_DATE,0 AS FLAG,ACTTA.CREATE_TIME,ACTTA.MODI_TIME,'P003' TRANS_TYPE,'Actb03'TRANS_NAME

                                    FROM [test0923].dbo.ACTTA,[test0923].dbo.ACTTB
                                    WHERE TA001=TB001 AND TA002=TB002
                                    AND (ISNULL(TB009,'')<>'')
                                    AND ISNULL(TB024,'')=''
                                    AND TA001=@TA001 AND TA002=@TA002
                                    ) AS TEMP
                                    WHERE REPLACE(MN005+MN006+MN007,' ','') NOT IN (SELECT REPLACE(MN005+MN006+MN007,' ','') FROM [test0923].dbo.ACTMN)
                                        ");


            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TA001", SqlDbType.NVarChar).Value = TA001;
                    command.Parameters.Add("@TA002", SqlDbType.NVarChar).Value = TA002;

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
