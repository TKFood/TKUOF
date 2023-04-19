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

namespace TKUOF.EXTERNALFORMFLOWS.GF05.ASSERT1001
{
    class ADDUSERSITES : ICallExternalDllSites
    {
        string MAINconnectionString = ConfigurationManager.ConnectionStrings["connectionstringUOF"].ToString(); 
         
        Lib.WKF.ExternalDllSites sites = new Lib.WKF.ExternalDllSites();
        UserUCO UserUCOSuperior = new UserUCO();
        EBUser EBUserSuperior;
        UserSet userSet1 = new UserSet();

      
        string account;
        string userGuid;
        string RANKS;
        string USER_GROUP_ID;
        string USER_GROUP_NAME;
        string USER_TITLE_NAME;
        string USER_APPLY_RANKS;

        public void Finally()
        {
            
        }

        public string GetExternalDllSites(string formInfo)
        {
            Ede.Uof.Utility.Log.Logger.Write("應用程式站點", "ASSERT1001 ADDUSERSITES 進入GetExternalDllSites " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            

            XmlDocument xmlDoc = new XmlDocument();
            XmlDocument formXmlDoc = new XmlDocument();
            DatabaseHelper DbQueryCompanyTopAccount = new DatabaseHelper();
            DatabaseHelper DbQueryGroupName = new DatabaseHelper();
            UserUCO userUCO = new UserUCO();
            DataSet GroupName = new DataSet();

            //找出表單新保管人(無法使用UOF者)、新保管人(無法使用UOF者)的部門
            //GAFrm000001NOU
            formXmlDoc.LoadXml(formInfo);
            //account = formXmlDoc.SelectSingleNode("/ExternalFlowSite/ApplicantInfo").Attributes["account"].Value;
            account = formXmlDoc.SelectSingleNode("/ExternalFlowSite/FormFieldValue/FieldItem[@fieldId='GAFrm000001NOU']").Attributes["fieldValue"].Value;
            account = account.Substring(4,6);

            DataTable DTACCOUNT = SEARCHACCOUNT(account);
            userGuid = DTACCOUNT.Rows[0]["USER_GUID"].ToString();
            USER_GROUP_ID = DTACCOUNT.Rows[0]["GROUP_ID"].ToString();
            USER_GROUP_NAME = DTACCOUNT.Rows[0]["GROUP_NAME"].ToString();
            USER_TITLE_NAME = DTACCOUNT.Rows[0]["TITLE_NAME"].ToString();
            USER_APPLY_RANKS = DTACCOUNT.Rows[0]["RANK"].ToString();
           
            

            //找出所有簽核人員，包含主管
            //FINDALLSINGER(userGuid);

            //測試用
            FINDTEST();

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

 

        public void FINDTEST()
        {
            UserUCO userUCO = new UserUCO();
            EBUser ebUser = userUCO.GetEBUser(userGuid);
            EBUser ebUserHasJobFunction = userUCO.GetEBUser(userGuid);

            Lib.WKF.ExternalDllSite site1 = new Lib.WKF.ExternalDllSite();
            site1.SignType = Lib.WKF.SignType.And;
           
           
     
            site1.Signers.Add("iteng");
            //site2.Signers.Add("160115");
            site1.Alerts.Add("160115");



            Ede.Uof.Utility.Log.Logger.Write("應用程式站點", "ASSERT1001 ADDUSERSITES 新增簽核 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));


            //site1 有找到簽核人員才新增簽核
            if (site1.Signers.Count > 0)
            {
                sites.Sites.Add(site1);
            }
          
        }

  
    }
}
