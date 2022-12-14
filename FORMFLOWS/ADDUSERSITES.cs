using Ede.Uof.EIP.Organization.Util;
using Ede.Uof.Utility.Data;
using Ede.Uof.WKF.CustomExternal;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Ede.Uof.Utility.Log;

namespace TKUOF.FORMFLOWS
{
    class ADDUSERSITES : ICallExternalDllSites
    {
        string MAINconnectionString = ConfigurationManager.ConnectionStrings["connectionstringUOF"].ToString(); 

        Lib.WKF.ExternalDllSites sites = new Lib.WKF.ExternalDllSites();
        UserUCO UserUCOSuperior = new UserUCO();
        EBUser EBUserSuperior;
        UserSet userSet1 = new UserSet();

        Boolean FLAGGO = true;
        DataSet CompanyTopAccountDS = new DataSet();
        DataTable DTACCOUNT = null;
        DataTable DTZ_UOF_FORM_DEP_SINGERS_DETAILS = null;

        string CompanyTopAccount;
        string SpecialGroupName = "N";
        string FORM_VERSION_ID;
        string UOF_FORM_NAME;
        string account;
        string userGuid;
        string RANKS;
        string GROUP_ID;
        string GROUP_NAME;
        string TITLE_NAME;
        string APPLY_RANKS;

        public void Finally()
        {
            
        }

        public string GetExternalDllSites(string formInfo)
        {
            Ede.Uof.Utility.Log.Logger.Write("應用程式站點", "ADDUSERSITES 進入GetExternalDllSites " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            

            XmlDocument xmlDoc = new XmlDocument();
            XmlDocument formXmlDoc = new XmlDocument();
            DatabaseHelper DbQueryCompanyTopAccount = new DatabaseHelper();
            DatabaseHelper DbQueryGroupName = new DatabaseHelper();
            UserUCO userUCO = new UserUCO();
            DataSet GroupName = new DataSet();

            //找出表單申請人、申請部門
            formXmlDoc.LoadXml(formInfo);
            account = formXmlDoc.SelectSingleNode("/ExternalFlowSite/ApplicantInfo").Attributes["account"].Value;

            DataTable DTACCOUNT = SEARCHACCOUNT(account);
            userGuid = DTACCOUNT.Rows[0]["USER_GUID"].ToString();
            GROUP_ID = DTACCOUNT.Rows[0]["GROUP_ID"].ToString();
            GROUP_NAME = DTACCOUNT.Rows[0]["GROUP_NAME"].ToString();
            TITLE_NAME = DTACCOUNT.Rows[0]["TITLE_NAME"].ToString();
            APPLY_RANKS = DTACCOUNT.Rows[0]["RANK"].ToString();
           
            //GROUP_ID = "0a700146-6015-4cc6-8aca-055a45e6a766";

            //找出表單 FORM_VERSION_ID、UOF_FORM_NAME
            FORM_VERSION_ID = formXmlDoc.SelectSingleNode("/ExternalFlowSite/ApplicantInfo").Attributes["formVersionId"].Value;
            UOF_FORM_NAME = SEARCHFORM_UOF_FORM_NAME(FORM_VERSION_ID);

            //先找出表單預設的簽核職級=RANK
            RANKS = Z_UOF_FORM_DEFALUT_SINGERS(UOF_FORM_NAME);
            //再找出表單是否有明細條件設定後相關簽核職級，如果有符合京明細條件就更改預設RANK



            ////檢查是否該申請的職級是否有明細的欄位條件設定
            ////如果有就照明細的欄位條件設定
            ////如果沒有明細的欄位條件設定，就依主要的申腈表單+申請部門+申請職級做簽核的決定
            ////RANKS = SEARCHFORM_UOF_Z_UOF_FORM_DEP_SINGERS(UOF_FORM_NAME, GROUP_ID, APPLY_RANKS);

            //DTZ_UOF_FORM_DEP_SINGERS_DETAILS = SEARCHZ_UOF_FORM_DEP_SINGERS_DETAILS(UOF_FORM_NAME, GROUP_ID, APPLY_RANKS);
            //if (DTZ_UOF_FORM_DEP_SINGERS_DETAILS != null && DTZ_UOF_FORM_DEP_SINGERS_DETAILS.Rows.Count > 0)
            //{
            //    //如果有就照明細的欄位條件設定
            //    RANKS = SEARCHFORM_UOF_FORM_DEP_SINGERS_DETAILS(UOF_FORM_NAME, GROUP_ID, APPLY_RANKS, formXmlDoc);
            //}

            ////如果沒有明細的欄位條件設定，就依主要的申腈表單+申請部門+申請職級做簽核的決定
            //if (string.IsNullOrEmpty(RANKS))
            //{
            //    //FORM_VERSION_ID，找出表單最高簽核人層級
            //    RANKS = SEARCHFORM_UOF_Z_UOF_FORM_DEP_SINGERS(UOF_FORM_NAME, GROUP_ID, APPLY_RANKS);
            //}



            //RANKS不得空白
            //找出部門所有簽核人員 依職級順序           
            if (!string.IsNullOrEmpty(RANKS))
            {
                FIND_FORM_FLOW_SINGER(GROUP_ID, RANKS);
            }



            //找出所有簽核人員，包含主管
            //FINDALLSINGER(userGuid);

            //測試用
            //FINDTEST();

            return sites.ConvertToXML();
        }

        public void OnError(Exception errorException)
        {

        }

        public DataTable SEARCHACCOUNT(string account)
        {
            string connectionString = MAINconnectionString;
            Ede.Uof.Utility.Data.DatabaseHelper m_db = new Ede.Uof.Utility.Data.DatabaseHelper(connectionString);
            StringBuilder cmdTxt = new StringBuilder();

            cmdTxt.AppendFormat(@" 
                           
                                SELECT 
                                [TB_EB_USER].[USER_GUID]
                                ,[ACCOUNT]
                                ,[NAME]
                                ,[TB_EB_GROUP].[GROUP_ID]
                                ,[GROUP_NAME]
                                ,[TB_EB_EMPL_DEP].[TITLE_ID]
                                ,[TITLE_NAME]
                                ,[RANK]
                                FROM [UOF].[dbo].[TB_EB_USER],[UOF].[dbo].[TB_EB_EMPL_DEP],[UOF].[dbo].[TB_EB_GROUP],[UOF].[dbo].[TB_EB_JOB_TITLE]
                                WHERE [TB_EB_USER].USER_GUID=[TB_EB_EMPL_DEP].USER_GUID
                                AND [TB_EB_GROUP].GROUP_ID=[TB_EB_EMPL_DEP].GROUP_ID
                                AND [UOF].[dbo].[TB_EB_JOB_TITLE].TITLE_ID=[TB_EB_EMPL_DEP].TITLE_ID
                                AND [TB_EB_USER].ACCOUNT='{0}'

                                 ", account);

          


            DataTable dt = new DataTable();

            dt.Load(m_db.ExecuteReader(cmdTxt.ToString()));

         

            if (dt.Rows.Count > 0)
            {
                return dt;
            }
            else
            {
                return null;
            }

           
        }

        public string Z_UOF_FORM_DEFALUT_SINGERS(string UOF_FORM_NAME)
        {
            string connectionString = MAINconnectionString;
            Ede.Uof.Utility.Data.DatabaseHelper m_db = new Ede.Uof.Utility.Data.DatabaseHelper(connectionString);
            StringBuilder cmdTxt = new StringBuilder();

            cmdTxt.AppendFormat(@" 
                                 SELECT TOP 1
                                [ID]
                                ,[UOF_FORM_NAME]
                                ,[RANKS]
                                ,[TITLE_NAME]
                                FROM [UOF].[dbo].[Z_UOF_FORM_DEFALUT_SINGERS]	
                                WHERE [UOF_FORM_NAME]='{0}'

                                 ", UOF_FORM_NAME);




            DataTable dt = new DataTable();

            dt.Load(m_db.ExecuteReader(cmdTxt.ToString()));



            if (dt.Rows.Count > 0)
            {
                return dt.Rows[0]["RANKS"].ToString();
            }
            else
            {
                return null;
            }

        }

        public DataTable SEARCHZ_UOF_FORM_DEP_SINGERS_DETAILS(string UOF_FORM_NAME, string GROUP_ID, string APPLY_RANKS)
        {
            string connectionString = MAINconnectionString;
            Ede.Uof.Utility.Data.DatabaseHelper m_db = new Ede.Uof.Utility.Data.DatabaseHelper(connectionString);
            StringBuilder cmdTxt = new StringBuilder();

            cmdTxt.AppendFormat(@"                            
                                SELECT 
                                TEMP.[ID]
                                ,[UOF_FORM_NAME]
                                ,[GROUP_ID]
                                ,[GROUP_NAME]
                                ,[RANKS]
                                ,[TITLE_NAME]
                                ,[APPLY_RANKS]
                                ,[APPLY_TITLE_NAME]
                                ,[PRIORITYS]

                                ,[MID]
                                ,[FIELDS]
                                ,[OPERATOR]
                                ,[CONDTIONVALUES]
                                ,[DETAILS_RANKS]
                                ,[DETAILS_TITLE_NAME]
                                ,[DETAILS_PRIORITYS]

                                FROM 
                                (
                                SELECT TOP (1) 
                                [ID]
                                ,[UOF_FORM_NAME]
                                ,[GROUP_ID]
                                ,[GROUP_NAME]
                                ,[RANKS]
                                ,[TITLE_NAME]
                                ,[APPLY_RANKS]
                                ,[APPLY_TITLE_NAME]
                                ,[PRIORITYS]
                                ,([APPLY_RANKS]-{2}) AS 'SEQ'
                                FROM [UOF].[dbo].[Z_UOF_FORM_DEP_SINGERS]
                                WHERE UOF_FORM_NAME='{0}' AND [GROUP_ID]='{1}' AND [APPLY_RANKS]>={2} 
                                ORDER BY ([APPLY_RANKS]-{2})
                                ) AS TEMP,[UOF].[dbo].[Z_UOF_FORM_DEP_SINGERS_DETAILS] 
                                WHERE 1=1
                                AND  TEMP.ID=[Z_UOF_FORM_DEP_SINGERS_DETAILS].MID

                                 ", UOF_FORM_NAME, GROUP_ID, APPLY_RANKS);




            DataTable dt = new DataTable();

            dt.Load(m_db.ExecuteReader(cmdTxt.ToString()));



            if (dt.Rows.Count > 0)
            {
                return dt;
            }
            else
            {
                return null;
            }


        }

        public void FINDTEST()
        {
            UserUCO userUCO = new UserUCO();
            EBUser ebUser = userUCO.GetEBUser(userGuid);
            EBUser ebUserHasJobFunction = userUCO.GetEBUser(userGuid);

            Lib.WKF.ExternalDllSite site1 = new Lib.WKF.ExternalDllSite();
            site1.SignType = Lib.WKF.SignType.And;
            Lib.WKF.ExternalDllSite site2 = new Lib.WKF.ExternalDllSite();
            site2.SignType = Lib.WKF.SignType.And;
            Lib.WKF.ExternalDllSite site3 = new Lib.WKF.ExternalDllSite();
            site3.SignType = Lib.WKF.SignType.And;
            Lib.WKF.ExternalDllSite site4 = new Lib.WKF.ExternalDllSite();

     
            site1.Signers.Add("120002");
            site2.Signers.Add("160115");
            site3.Signers.Add("iteng");
             
            Ede.Uof.Utility.Log.Logger.Write("應用程式站點", "ADDUSERSITES 新增簽核 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));


            //site1 有找到簽核人員才新增簽核
            if (site1.Signers.Count > 0)
            {
                sites.Sites.Add(site1);
            }
            //site2 有找到簽核人員才新增簽核
            if (site2.Signers.Count > 0)
            {
                sites.Sites.Add(site2);
            }
            //site3 有找到簽核人員才新增簽核
            if (site3.Signers.Count > 0)
            {
                sites.Sites.Add(site3);
            }
            //site4 有找到簽核人員才新增簽核
            if (site4.Signers.Count > 0)
            {
                sites.Sites.Add(site4);
            }



        }

        public void FIND_FORM_FLOW_SINGER(string GROUP_ID, string RANKS)
        {
            DataTable DTSITESATFERRANKS = new DataTable();
            DTSITESATFERRANKS.Clear();
            DTSITESATFERRANKS.Columns.Add("ACCOUNT");
            string ADDRANK = "Y";
            int FORMRANKS = 0;

            UserUCO userUCO = new UserUCO();
            EBUser ebUser = userUCO.GetEBUser(userGuid);
            EBUser ebUserHasJobFunction = userUCO.GetEBUser(userGuid);

            Lib.WKF.ExternalDllSite site1 = new Lib.WKF.ExternalDllSite();
            site1.SignType = Lib.WKF.SignType.And;
            Lib.WKF.ExternalDllSite site2 = new Lib.WKF.ExternalDllSite();
            site2.SignType = Lib.WKF.SignType.And;
            Lib.WKF.ExternalDllSite site3= new Lib.WKF.ExternalDllSite();
            site3.SignType = Lib.WKF.SignType.And;
            Lib.WKF.ExternalDllSite site4 = new Lib.WKF.ExternalDllSite();
            site4.SignType = Lib.WKF.SignType.And;
            Lib.WKF.ExternalDllSite site5 = new Lib.WKF.ExternalDllSite();
            site5.SignType = Lib.WKF.SignType.And;
            Lib.WKF.ExternalDllSite site6 = new Lib.WKF.ExternalDllSite();
            site6.SignType = Lib.WKF.SignType.And;
            Lib.WKF.ExternalDllSite site7 = new Lib.WKF.ExternalDllSite();
            site7.SignType = Lib.WKF.SignType.And;
            Lib.WKF.ExternalDllSite site8 = new Lib.WKF.ExternalDllSite();
            site8.SignType = Lib.WKF.SignType.And;
            Lib.WKF.ExternalDllSite site9 = new Lib.WKF.ExternalDllSite();
            site9.SignType = Lib.WKF.SignType.And;
            Lib.WKF.ExternalDllSite site10 = new Lib.WKF.ExternalDllSite();


            //找出部門糐上層的所有簽核者資料
            //DataTable DTSITES = SEARCH_FORM_FLOW_SITES(GROUP_ID, RANKS);
            DataTable DTSITES = SEARCH_FORM_FLOW_SITES_ALL(GROUP_ID);

            //找職級小於/等於的簽核人數
            foreach(DataRow DR in DTSITES.Rows)
            {
                if(Convert.ToInt32(RANKS)<= Convert.ToInt32(DR["RANK"].ToString()))
                {
                    FORMRANKS = FORMRANKS+1;
                }
            }
            //檢查所有職級是否相符條件的職級，有相同職級就不再找上一層
            foreach (DataRow DR in DTSITES.Rows)
            {
                if (Convert.ToInt32(RANKS) == Convert.ToInt32(DR["RANK"].ToString()))
                {
                    ADDRANK = "N";
                }
            }
            //檢查所有職級是否相符條件的職級，沒有相同職級就再找上一層
            if (ADDRANK.Equals("Y"))
            {
                FORMRANKS = FORMRANKS + 1;
            }

            for(int i=0;i< FORMRANKS;i++)
            {
                DataRow NEWDR = DTSITESATFERRANKS.NewRow();
                NEWDR["ACCOUNT"] = DTSITES.Rows[i]["ACCOUNT"].ToString();
                DTSITESATFERRANKS.Rows.Add(NEWDR);
            }

            if (DTSITESATFERRANKS.Rows.Count >= 1)
            {
                site1.Signers.Add(DTSITESATFERRANKS.Rows[0]["ACCOUNT"].ToString());
            }
            if (DTSITESATFERRANKS.Rows.Count >= 2)
            {
                site2.Signers.Add(DTSITESATFERRANKS.Rows[1]["ACCOUNT"].ToString());
            }
            if (DTSITESATFERRANKS.Rows.Count >= 3)
            {
                site3.Signers.Add(DTSITESATFERRANKS.Rows[2]["ACCOUNT"].ToString());
            }
            if (DTSITESATFERRANKS.Rows.Count >= 4)
            {
                site4.Signers.Add(DTSITESATFERRANKS.Rows[3]["ACCOUNT"].ToString());
            }
            if (DTSITESATFERRANKS.Rows.Count >= 5)
            {
                site5.Signers.Add(DTSITESATFERRANKS.Rows[4]["ACCOUNT"].ToString());
            }
            if (DTSITESATFERRANKS.Rows.Count >= 6)
            {
                site6.Signers.Add(DTSITESATFERRANKS.Rows[5]["ACCOUNT"].ToString());
            }
            if (DTSITESATFERRANKS.Rows.Count >= 7)
            {
                site7.Signers.Add(DTSITESATFERRANKS.Rows[6]["ACCOUNT"].ToString());
            }
            if (DTSITESATFERRANKS.Rows.Count >= 8)
            {
                site8.Signers.Add(DTSITESATFERRANKS.Rows[7]["ACCOUNT"].ToString());
            }
            if (DTSITESATFERRANKS.Rows.Count >= 9)
            {
                site9.Signers.Add(DTSITESATFERRANKS.Rows[8]["ACCOUNT"].ToString());
            }
            if (DTSITESATFERRANKS.Rows.Count >= 10)
            {
                site10.Signers.Add(DTSITESATFERRANKS.Rows[9]["ACCOUNT"].ToString());
            }
          

            //if (DTSITES.Rows.Count >= 1)
            //{
            //    site1.Signers.Add(DTSITES.Rows[0]["ACCOUNT"].ToString());
            //}
            //if (DTSITES.Rows.Count >= 2)
            //{
            //    site2.Signers.Add(DTSITES.Rows[1]["ACCOUNT"].ToString());
            //}
            //if (DTSITES.Rows.Count >= 3)
            //{
            //    site3.Signers.Add(DTSITES.Rows[2]["ACCOUNT"].ToString());
            //}
            //if (DTSITES.Rows.Count >= 4)
            //{
            //    site4.Signers.Add(DTSITES.Rows[3]["ACCOUNT"].ToString());
            //}


            //site1 有找到簽核人員才新增簽核
            if (site1.Signers.Count > 0)
            {
                sites.Sites.Add(site1);
            }
            //site2 有找到簽核人員才新增簽核
            if (site2.Signers.Count > 0)
            {
                sites.Sites.Add(site2);
            }
            //site3 有找到簽核人員才新增簽核
            if (site3.Signers.Count > 0)
            {
                sites.Sites.Add(site3);
            }
            //site4 有找到簽核人員才新增簽核
            if (site4.Signers.Count > 0)
            {
                sites.Sites.Add(site4);
            }
            //site5 有找到簽核人員才新增簽核
            if (site5.Signers.Count > 0)
            {
                sites.Sites.Add(site5);
            }
            //site6 有找到簽核人員才新增簽核
            if (site6.Signers.Count > 0)
            {
                sites.Sites.Add(site6);
            }
            //site7 有找到簽核人員才新增簽核
            if (site7.Signers.Count > 0)
            {
                sites.Sites.Add(site7);
            }
            //site8 有找到簽核人員才新增簽核
            if (site8.Signers.Count > 0)
            {
                sites.Sites.Add(site8);
            }
            //site9 有找到簽核人員才新增簽核
            if (site9.Signers.Count > 0)
            {
                sites.Sites.Add(site9);
            }
            //site10 有找到簽核人員才新增簽核
            if (site10.Signers.Count > 0)
            {
                sites.Sites.Add(site10);
            }



        }
        public DataTable SEARCH_FORM_FLOW_SITES_ALL(string GROUP_ID)
        {

            string connectionString = MAINconnectionString;
            Ede.Uof.Utility.Data.DatabaseHelper m_db = new Ede.Uof.Utility.Data.DatabaseHelper(connectionString);
            StringBuilder cmdTxt = new StringBuilder();

            cmdTxt.AppendFormat(@" 
                            WITH CTE_TB_EB_GROUP AS 
                            (
                            SELECT 
                            GROUP_NAME, GROUP_ID, PARENT_GROUP_ID, LEV, 1 AS LEVELS
                            FROM [UOF].dbo.TB_EB_GROUP
                            WHERE (GROUP_ID = '{0}')
                            UNION ALL
                            SELECT A.GROUP_NAME, A.GROUP_ID, A.PARENT_GROUP_ID, A.LEV, 
                            B.LEVELS + 1 AS Expr1
                            FROM [UOF].dbo.TB_EB_GROUP AS A INNER JOIN
                            CTE_TB_EB_GROUP AS B ON B.PARENT_GROUP_ID = A.GROUP_ID

                            )

                            SELECT CTE_TB_EB_GROUP.GROUP_NAME, CTE_TB_EB_GROUP.GROUP_ID, CTE_TB_EB_GROUP.PARENT_GROUP_ID, CTE_TB_EB_GROUP.LEV, LEVELS
                            ,TB_EB_EMPL_FUNC.*
                            ,TB_EB_JOB_FUNC.*
                            ,TB_EB_USER.ACCOUNT
                            ,TB_EB_USER.NAME
                            ,TB_EB_EMPL_DEP.*
                            ,TB_EB_JOB_TITLE.*
                            FROM  CTE_TB_EB_GROUP,[UOF].dbo.TB_EB_EMPL_FUNC, [UOF].dbo.TB_EB_JOB_FUNC ,[UOF].dbo.TB_EB_USER,[UOF].dbo.TB_EB_EMPL_DEP,[UOF].dbo.TB_EB_JOB_TITLE 
                            WHERE  1=1
                            AND CTE_TB_EB_GROUP.GROUP_ID=TB_EB_EMPL_FUNC.GROUP_ID
                            AND TB_EB_EMPL_FUNC.FUNC_ID=TB_EB_JOB_FUNC.FUNC_ID
                            AND TB_EB_EMPL_FUNC.USER_GUID=TB_EB_USER.USER_GUID
                            AND TB_EB_EMPL_DEP.USER_GUID=TB_EB_USER.USER_GUID
                            AND TB_EB_EMPL_DEP.GROUP_ID=CTE_TB_EB_GROUP.GROUP_ID
                            AND TB_EB_JOB_TITLE.TITLE_ID=TB_EB_EMPL_DEP.TITLE_ID
                            AND (CTE_TB_EB_GROUP.GROUP_NAME NOT LIKE '%停用%') AND (CTE_TB_EB_GROUP.GROUP_NAME NOT LIKE '%特殊用途%')
                            AND TB_EB_EMPL_FUNC.FUNC_ID IN ('Signer')
                            AND TB_EB_USER.IS_SUSPENDED IN ('0')
                          
                            ORDER BY  LEV DESC,RANK 

                             ", GROUP_ID);

            //m_db.AddParameter("@GROUP_ID", GROUP_ID);


            DataTable dt = new DataTable();

            dt.Load(m_db.ExecuteReader(cmdTxt.ToString()));

            if (dt.Rows.Count > 0)
            {
                return dt;
            }
            else
            {
                return null;
            }
        }

        public DataTable SEARCH_FORM_FLOW_SITES(string GROUP_ID,string RANKS)
        {

            string connectionString = MAINconnectionString;
            Ede.Uof.Utility.Data.DatabaseHelper m_db = new Ede.Uof.Utility.Data.DatabaseHelper(connectionString);
            StringBuilder cmdTxt = new StringBuilder();

            cmdTxt.AppendFormat(@" 
                            WITH CTE_TB_EB_GROUP AS 
                            (
                            SELECT 
                            GROUP_NAME, GROUP_ID, PARENT_GROUP_ID, LEV, 1 AS LEVELS
                            FROM [UOF].dbo.TB_EB_GROUP
                            WHERE (GROUP_ID = '{0}')
                            UNION ALL
                            SELECT A.GROUP_NAME, A.GROUP_ID, A.PARENT_GROUP_ID, A.LEV, 
                            B.LEVELS + 1 AS Expr1
                            FROM [UOF].dbo.TB_EB_GROUP AS A INNER JOIN
                            CTE_TB_EB_GROUP AS B ON B.PARENT_GROUP_ID = A.GROUP_ID

                            )

                            SELECT CTE_TB_EB_GROUP.GROUP_NAME, CTE_TB_EB_GROUP.GROUP_ID, CTE_TB_EB_GROUP.PARENT_GROUP_ID, CTE_TB_EB_GROUP.LEV, LEVELS
                            ,TB_EB_EMPL_FUNC.*
                            ,TB_EB_JOB_FUNC.*
                            ,TB_EB_USER.ACCOUNT
                            ,TB_EB_USER.NAME
                            ,TB_EB_EMPL_DEP.*
                            ,TB_EB_JOB_TITLE.*
                            FROM  CTE_TB_EB_GROUP,[UOF].dbo.TB_EB_EMPL_FUNC, [UOF].dbo.TB_EB_JOB_FUNC ,[UOF].dbo.TB_EB_USER,[UOF].dbo.TB_EB_EMPL_DEP,[UOF].dbo.TB_EB_JOB_TITLE 
                            WHERE  1=1
                            AND CTE_TB_EB_GROUP.GROUP_ID=TB_EB_EMPL_FUNC.GROUP_ID
                            AND TB_EB_EMPL_FUNC.FUNC_ID=TB_EB_JOB_FUNC.FUNC_ID
                            AND TB_EB_EMPL_FUNC.USER_GUID=TB_EB_USER.USER_GUID
                            AND TB_EB_EMPL_DEP.USER_GUID=TB_EB_USER.USER_GUID
                            AND TB_EB_EMPL_DEP.GROUP_ID=CTE_TB_EB_GROUP.GROUP_ID
                            AND TB_EB_JOB_TITLE.TITLE_ID=TB_EB_EMPL_DEP.TITLE_ID
                            AND (CTE_TB_EB_GROUP.GROUP_NAME NOT LIKE '%停用%') AND (CTE_TB_EB_GROUP.GROUP_NAME NOT LIKE '%特殊用途%')
                            AND TB_EB_EMPL_FUNC.FUNC_ID IN ('Signer')
                            AND TB_EB_USER.IS_SUSPENDED IN ('0')
                            AND TB_EB_JOB_TITLE.RANK>={1}

                            ORDER BY  LEV DESC,RANK 

                             ", GROUP_ID, RANKS);

            //m_db.AddParameter("@GROUP_ID", GROUP_ID);


            DataTable dt = new DataTable();

            dt.Load(m_db.ExecuteReader(cmdTxt.ToString()));

            if (dt.Rows.Count > 0)
            {
                return dt;
            }
            else
            {
                return null;
            }
        }

        public void FINDALLSINGER(string userGuid)
        {
            UserUCO userUCO = new UserUCO();
            EBUser ebUser = userUCO.GetEBUser(userGuid);
            EBUser ebUserHasJobFunction = userUCO.GetEBUser(userGuid);

            Lib.WKF.ExternalDllSite site1 = new Lib.WKF.ExternalDllSite();
            site1.SignType = Lib.WKF.SignType.And;
            Lib.WKF.ExternalDllSite site2 = new Lib.WKF.ExternalDllSite();

            //SEARCHDEPSITES
            DataTable DTSITES = SEARCHDEPSITES(ebUser.Account);
            
            if(DTSITES.Rows.Count>=1)
            {
                site1.Signers.Add(DTSITES.Rows[0]["UOFFLOWSDEPMANAGERS_ACCOUNT"].ToString());
            }
            if (DTSITES.Rows.Count >= 2)
            {
                site2.Signers.Add(DTSITES.Rows[1]["UOFFLOWSDEPMANAGERS_ACCOUNT"].ToString());
            }


            //site1 有找到簽核人員才新增簽核
            if (site1.Signers.Count > 0)
            {
                sites.Sites.Add(site1);
            }

            //site2 有找到簽核人員才新增簽核
            if (site2.Signers.Count > 0)
            {
                sites.Sites.Add(site2);
            }


        }

        public DataTable SEARCHDEPSITES(string ACCOUNT)
        {

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();
            Ede.Uof.Utility.Data.DatabaseHelper m_db = new Ede.Uof.Utility.Data.DatabaseHelper(connectionString);

            string cmdTxt = @"                             
                             SELECT 
                            [UOFFLOWSUSERS].[ACCOUNT] UOFFLOWSUSERS_ACCOUNT
                            ,[UOFFLOWSUSERS].[NAMES] UOFFLOWSUSERS_NAMES
                            ,[UOFFLOWSUSERS].[DEPNAME] UOFFLOWSUSERS_DEPNAME
                            ,[UOFFLOWSDEPMANAGERS].[ACCOUNT] UOFFLOWSDEPMANAGERS_ACCOUNT
                            ,[UOFFLOWSDEPMANAGERS].[NAMES] UOFFLOWSDEPMANAGERS_NAMES
                            ,[UOFFLOWSDEPMANAGERS].[DEPNAME] UOFFLOWSDEPMANAGERS_DEPNAME
                            ,[UOFFLOWSDEPMANAGERS].[ISMANAGER] UOFFLOWSDEPMANAGERS_ISMANAGER
                            ,[UOFFLOWSDEPMANAGERS].[FLOWSSEQ] UOFFLOWSDEPMANAGERS_FLOWSSEQ

                            FROM [TKIT].[dbo].[UOFFLOWSUSERS]
                            LEFT JOIN [TKIT].[dbo].[UOFFLOWSDEPMANAGERS] ON [UOFFLOWSDEPMANAGERS].DEPNAME=[UOFFLOWSUSERS].DEPNAME AND [UOFFLOWSDEPMANAGERS].[ISMANAGER]='Y'
                            WHERE [UOFFLOWSUSERS].ACCOUNT=@ACCOUNT
                            ORDER BY [UOFFLOWSDEPMANAGERS].[FLOWSSEQ]
                             ";

            m_db.AddParameter("@ACCOUNT", ACCOUNT);


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

        public string SEARCHFORM_UOF_FORM_NAME(string FORM_VERSION_ID)
        {
            string connectionString = MAINconnectionString;
            SqlConnection sqlConn = new SqlConnection(connectionString);
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            StringBuilder queryString = new StringBuilder();
            DataSet ds = new DataSet();

            try
            {
                //要記得改成正式-3
                queryString.AppendFormat(@" 
                                                
                                        SELECT TOP 1 RTRIM(LTRIM(TB_WKF_FORM_VERSION.FORM_VERSION_ID)) FORM_VERSION_ID,TB_WKF_FORM_VERSION.FORM_ID,TB_WKF_FORM_VERSION.VERSION,TB_WKF_FORM_VERSION.ISSUE_CTL
                                        ,TB_WKF_FORM.FORM_NAME
                                        FROM [UOF].dbo.TB_WKF_FORM_VERSION,[UOF].dbo.TB_WKF_FORM
                                        WHERE 1=1
                                        AND TB_WKF_FORM_VERSION.FORM_ID=TB_WKF_FORM.FORM_ID
                                        AND TB_WKF_FORM_VERSION.ISSUE_CTL=1
                                        AND FORM_VERSION_ID='{0}'
                                        ORDER BY TB_WKF_FORM_VERSION.FORM_ID,TB_WKF_FORM_VERSION.VERSION DESC

                                          ", FORM_VERSION_ID);

                adapter = new SqlDataAdapter(@"" + queryString, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    return ds.Tables["TEMPds1"].Rows[0]["FORM_NAME"].ToString();
                }
                else
                {
                    return "";
                }

            }
            catch
            {
                return "";
            }
            finally
            {

            }
        }

        public string SEARCHFORM_UOF_Z_UOF_FORM_DEP_SINGERS(string UOF_FORM_NAME, string GROUP_ID,string APPLY_RANKS)
        {
            string connectionString = MAINconnectionString;
            SqlConnection sqlConn = new SqlConnection(connectionString);
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            StringBuilder queryString = new StringBuilder();
            DataSet ds = new DataSet();

            try
            {
                //要記得改成正式-3
                //先找出申請人的職級最符合的核準職級
                //再用最符合的核準職級找是否有明細欄位的條件

                queryString.AppendFormat(@" 
                                        SELECT TOP (1) 
                                        [ID]
                                        ,[UOF_FORM_NAME]
                                        ,[GROUP_ID]
                                        ,[GROUP_NAME]
                                        ,[RANKS]
                                        ,[TITLE_NAME]
                                        ,[APPLY_RANKS]
                                        ,[APPLY_TITLE_NAME]
                                        ,[PRIORITYS]
                                        ,([APPLY_RANKS]-{2}) AS 'SEQ'
                                        FROM [UOF].[dbo].[Z_UOF_FORM_DEP_SINGERS]
                                        WHERE UOF_FORM_NAME='{0}' AND [GROUP_ID]='{1}' AND [APPLY_RANKS]>={2} 
                                        ORDER BY ([APPLY_RANKS]-{2})

                                          ", UOF_FORM_NAME, GROUP_ID, APPLY_RANKS);

                adapter = new SqlDataAdapter(@"" + queryString, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    return ds.Tables["TEMPds1"].Rows[0]["RANKS"].ToString();
                }
                else
                {
                    return "";
                }

            }
            catch
            {
                return "";
            }
            finally
            {

            }
        }

        public string SEARCHFORM_UOF_FORM_DEP_SINGERS_DETAILS(string UOF_FORM_NAME, string GROUP_ID, string APPLY_RANKS, XmlDocument formXmlDoc)
        {
            string connectionString = MAINconnectionString;
            Ede.Uof.Utility.Data.DatabaseHelper m_db = new Ede.Uof.Utility.Data.DatabaseHelper(connectionString);
            StringBuilder cmdTxt = new StringBuilder();

            cmdTxt.AppendFormat(@"                            
                                SELECT 
                                TEMP.[ID]
                                ,[UOF_FORM_NAME]
                                ,[GROUP_ID]
                                ,[GROUP_NAME]
                                ,[RANKS]
                                ,[TITLE_NAME]
                                ,[APPLY_RANKS]
                                ,[APPLY_TITLE_NAME]
                                ,[PRIORITYS]

                                ,[MID]
                                ,[FIELDS]
                                ,[OPERATOR]
                                ,[CONDTIONVALUES]
                                ,[DETAILS_RANKS]
                                ,[DETAILS_TITLE_NAME]
                                ,[DETAILS_PRIORITYS]

                                FROM 
                                (
                                SELECT TOP (1) 
                                [ID]
                                ,[UOF_FORM_NAME]
                                ,[GROUP_ID]
                                ,[GROUP_NAME]
                                ,[RANKS]
                                ,[TITLE_NAME]
                                ,[APPLY_RANKS]
                                ,[APPLY_TITLE_NAME]
                                ,[PRIORITYS]
                                ,([APPLY_RANKS]-{2}) AS 'SEQ'
                                FROM [UOF].[dbo].[Z_UOF_FORM_DEP_SINGERS]
                                WHERE UOF_FORM_NAME='{0}' AND [GROUP_ID]='{1}' AND [APPLY_RANKS]>={2} 
                                ORDER BY ([APPLY_RANKS]-{2})
                                ) AS TEMP,[UOF].[dbo].[Z_UOF_FORM_DEP_SINGERS_DETAILS] 
                                WHERE 1=1
                                AND  TEMP.ID=[Z_UOF_FORM_DEP_SINGERS_DETAILS].MID

                                 ", UOF_FORM_NAME, GROUP_ID, APPLY_RANKS);




            DataTable dt = new DataTable();

            dt.Load(m_db.ExecuteReader(cmdTxt.ToString()));



            if (dt != null &&  dt.Rows.Count > 0)
            {
                string RANK = null;
                RANK = FIND_Z_UOF_FORM_DEP_SINGERS_DETAILS(dt, formXmlDoc);

                return RANK;
            }
            else
            {
                return null;
            }

        }
        /// <summary>
        /// 把明細欄位的值全部找出來比對
        /// </summary>
        /// <returns></returns>
        public string FIND_Z_UOF_FORM_DEP_SINGERS_DETAILS(DataTable DT, XmlDocument formXmlDoc)
        {
            string FIANL_DETAILS_RANKS = null;

            if (DT != null && DT.Rows.Count > 0)
            {
                foreach(DataRow DR in DT.Rows)
                {
                    StringBuilder FINDXML = new StringBuilder();
                    string FIELDS = DR["FIELDS"].ToString();
                    string OPERATOR = DR["OPERATOR"].ToString();
                    string CONDTIONVALUES = DR["CONDTIONVALUES"].ToString();
                    string DETAILS_RANKS = DR["DETAILS_RANKS"].ToString();
                    string RANKS = DR["RANKS"].ToString();

                    FINDXML.AppendFormat(@"/ExternalFlowSite/FormFieldValue/FieldItem[@fieldId='{0}'] ", FIELDS);

                    string XMLVALUES = formXmlDoc.SelectSingleNode(FINDXML.ToString()).Attributes["fieldValue"].Value;
                    int CONDTIONS = string.Compare(XMLVALUES, CONDTIONVALUES);

                    if(CONDTIONS>0 && (OPERATOR.Equals(">=")))
                    {
                        FIANL_DETAILS_RANKS = DETAILS_RANKS;
                    }
                    else if(CONDTIONS < 0 && ( OPERATOR.Equals("<=")))
                    {
                        FIANL_DETAILS_RANKS = DETAILS_RANKS;
                    }
                    else if(CONDTIONS == 0 && (OPERATOR.Equals("==") || OPERATOR.Equals(">=") || OPERATOR.Equals("<=")))
                    {
                        FIANL_DETAILS_RANKS = DETAILS_RANKS;
                    }
                  
                }
               
                
                
                
                return FIANL_DETAILS_RANKS;
            }
            else
            {
                return null;
            }

           
        }

        //public string GetExternalDllSites(string formInfo)
        //{
        //    XmlDocument xmlDoc = new XmlDocument();
        //    XmlDocument formXmlDoc = new XmlDocument();
        //    DatabaseHelper DbQueryCompanyTopAccount = new DatabaseHelper();
        //    DatabaseHelper DbQueryGroupName = new DatabaseHelper();
        //    UserUCO userUCO = new UserUCO();
        //    DataSet GroupName = new DataSet();


        //    formXmlDoc.LoadXml(formInfo);
        //    account = formXmlDoc.SelectSingleNode("/ExternalFlowSite/ApplicantInfo").Attributes["account"].Value;
        //    userGuid = userUCO.GetGUID(account);
        //    EBUser ebUser = userUCO.GetEBUser(userGuid);
        //    DEPNAME = ebUser.GroupName;

        //    //找出公司內最高的主管account
        //    CompanyTopAccountDS = DbQueryCompanyTopAccount.ExecuteDataSet("SELECT TOP 1 CompanyTopAccount FROM [UOFTEST].[dbo].[CompanyTopAccount]");
        //    CompanyTopAccount= CompanyTopAccountDS.Tables[0].Rows[0]["CompanyTopAccount"].ToString();

        //    GroupName = DbQueryGroupName.ExecuteDataSet("SELECT  [GroupName]   FROM [UOFTEST].[dbo].[GROUPNAMETOEF]");
        //    //判斷是不是指定部門，可由簽至副理就結束
        //    foreach (DataRow dr in GroupName.Tables[0].Rows)
        //    {
        //        if (DEPNAME.Equals(dr["GroupName"].ToString()))
        //        {
        //            SpecialGroupName = "Y";
        //            break;
        //        }
        //        else
        //        {
        //            SpecialGroupName = "N";
        //        }
        //    }

        //    //找出所有簽核人員，包含主管
        //    FINDALLSINGER(userGuid);

        //    return sites.ConvertToXML();
        //}

        //public void FINDALLSINGER(string userGuid)
        //{
        //    UserUCO userUCO = new UserUCO();
        //    EBUser ebUser = userUCO.GetEBUser(userGuid);
        //    EBUser ebUserHasJobFunction = userUCO.GetEBUser(userGuid);

        //    //找到申請者的所有簽核者和主管
        //    while (FLAGGO==true)
        //    {
        //        BaseGroup baseGroup = new BaseGroup(ebUser.GroupID);

        //        //如果沒有上層的部門就停止
        //        if (baseGroup.ParnetGroup == null)
        //        {
        //            FLAGGO = false;
        //            //break;
        //        }
        //        else
        //        {
        //            Lib.WKF.ExternalDllSite site1 = new Lib.WKF.ExternalDllSite();
        //            site1.SignType = Lib.WKF.SignType.And;
        //            Lib.WKF.ExternalDllSite site2 = new Lib.WKF.ExternalDllSite();
        //            //loop start
        //            userSet1 = GetUserSinger(userGuid);

        //            if (userSet1.Items.Count > 0)
        //            {
        //                for (int i = 0; i < userSet1.Items.Count; i++)
        //                {
        //                    ebUserHasJobFunction = userUCO.GetEBUser(userSet1.Items[i].Key);

        //                    if (!ebUserHasJobFunction.HasJobFunction("Superior"))
        //                    {
        //                        //申請者不用再當簽核者
        //                        if(!account.Equals(ebUserHasJobFunction.Account))
        //                        {
        //                            site1.Signers.Add(ebUserHasJobFunction.Account);
        //                            //add site1 singer
        //                        }
        //                    }
        //                    else
        //                    {
        //                        site2.Signers.Add(ebUserHasJobFunction.Account);
        //                        //add site2 singer

        //                        EBUserSuperior = ebUserHasJobFunction;
        //                    }
        //                    //loop end
        //                }

        //                //有找到簽核人員才新增簽核
        //                if (site1.Signers.Count > 0)
        //                {
        //                    sites.Sites.Add(site1);
        //                }

        //                //有找到主管新增簽核
        //                if (site2.Signers.Count > 0)
        //                {
        //                    sites.Sites.Add(site2);
        //                }


        //            }

        //        }

        //        //如果沒有找到最高主管就繼續找上層部門的簽核人員+主管  
        //        if (!EBUserSuperior.Account.Equals(CompanyTopAccount))
        //        {
        //            //FINDALLSINGER(Superior);

        //            //主管的職級是理級以上就停止，不是就往上找
        //            if (SpecialGroupName.Equals("N") && EBUserSuperior.GetEmployeeDepartment(DepartmentOfUser.Major).JobTitle.Rank <= 7)
        //            {
        //                FLAGGO = false;
        //            }
        //            //如果指定部門的主管是副理以上就停止，不是就往上找
        //            else if (SpecialGroupName.Equals("Y") && EBUserSuperior.GetEmployeeDepartment(DepartmentOfUser.Major).JobTitle.Rank <= 9)
        //            {
        //                FLAGGO = false;
        //            }
        //            else
        //            {
        //                FINDALLSINGER(EBUserSuperior.UserGUID);
        //            }
        //        }
        //        //如果找到最高主管就停止            
        //        else
        //        {
        //            FLAGGO = false;
        //            //break;
        //        }

        //    }

        //}

        //public void OnError(Exception errorException)
        //{

        //}

        //// <summary>
        ///// 取得員工簽核人員
        ///// </summary>
        ///// <param name="userGuid"></param>
        ///// <returns></returns>
        //public UserSet GetUserSinger(string userGuid)
        //{
        //    UserSet userSet = new UserSet();
        //    UserUCO userUCO = new UserUCO();
        //    EBUser ebUser = userUCO.GetEBUser(userGuid);
        //    BaseGroup baseGp = ebUser.GetEmployeeDepartment(DepartmentOfUser.Major).Department;

        //    //找出所有簽核人員
        //    if (CheckIsDeptSigner(ebUser.UserGUID, baseGp.GroupId) == false)
        //    {
        //        AddSignerToUserSet(baseGp.GroupId, userSet);
        //    }
        //    else
        //    {
        //        //如果目前的簽核人員
        //        if (baseGp.ParnetGroup != null)
        //        {
        //            AddSignerToUserSet(baseGp.ParnetGroup.GroupId, userSet);
        //        }
        //    }




        //    return userSet;

        //}


        ///// <summary>
        ///// 檢查是否是否是主管
        ///// </summary>
        ///// <returns></returns>
        //private bool CheckIsDeptSigner(string userGUID, string groupId)
        //{
        //    UserUCO userUco = new UserUCO();
        //    bool IsSuperior = false;
        //    EBUser ebUser = userUco.GetEBUser(userGUID);
        //    EmployeeJobFunctionCollection employeeJobFunctionCollection = ebUser.GetJobFunctionsOfDepartment(groupId);

        //    //判斷是否是主管，
        //    for (int i = 0; i < employeeJobFunctionCollection.Count; i++)
        //    {
        //        if (employeeJobFunctionCollection[i].FunctionId == "Superior")
        //        {
        //            IsSuperior = true;
        //            break;
        //        }
        //    }

        //    return IsSuperior;
        //}


        ///// <summary>
        ///// 把簽核者加到 UserSet 裡 , 2006/11/27 尋找部門主管，如果找不到就往上找，直到找到為止
        ///// </summary>
        ///// <param name="groupId"></param>
        ///// <param name="userSet"></param>
        //private void AddSignerToUserSet(string groupId, UserSet userSet)
        //{
        //    EmployeeFindUCO employeeFindUCO = new EmployeeFindUCO();
        //    //取得查詢群組的簽核者
        //    UserSetEBUsersCollection userSetEBUsersCollection = employeeFindUCO.FindEmployeeByFunction(groupId, "Signer");

        //    if (userSetEBUsersCollection.Count > 0)
        //    {
        //        for (int i = 0; i < userSetEBUsersCollection.Count; i++)
        //        {
        //            //把查到的簽核者 UserGuid 加到 userSet 裡
        //            UserSetUser userSetUser = new UserSetUser();
        //            userSetUser.USER_GUID = userSetEBUsersCollection[i].UserGUID;
        //            userSet.Items.Add(userSetUser);
        //        }
        //    }
        //    else
        //    {
        //        //如果找不到直屬簽核者就往上一層層找上去，直到有為止
        //        BaseGroup baseGroup = new BaseGroup(groupId);
        //        if (baseGroup.ParnetGroup != null)
        //        {
        //            AddSignerToUserSet(baseGroup.ParnetGroup.GroupId, userSet);
        //        }
        //    }
        //}
    }
}
