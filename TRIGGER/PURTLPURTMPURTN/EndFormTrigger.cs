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

namespace TKUOF.TRIGGER.PURTLPURTMPURTN
{
    //請購單的核準

    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {
            
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TL001=null;
            string TL002 = null;
          
            string FORMID= null;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TL001 = applyTask.Task.CurrentDocument.Fields["TL001"].FieldValue.ToString().Trim();
            TL002 = applyTask.Task.CurrentDocument.Fields["TL002"].FieldValue.ToString().Trim();
         
            FORMID = applyTask.FormNumber;

            ///核準
            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                if (!string.IsNullOrEmpty(TL001) && !string.IsNullOrEmpty(TL002) )
                {
                    UPDATEPURMBPURMC(TL001, TL002, FORMID);
                }
            }
            //作廢
            else if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Cancel)
            {
                //if (!string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002))
                //{
                //    UPDATEPURTCPURTDCANCEL(TC001, TC002, FORMID);
                //}
            }
            //否決
            else if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Reject)
            {
            //    if (!string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002))
            //    {
            //        UPDATEPURTCPURTDREJECT(TC001, TC002, FORMID);
            //    }
            }
            //退簽
            else if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Return)
            {
                //if (!string.IsNullOrEmpty(TC001) && !string.IsNullOrEmpty(TC002))
                //{
                //    UPDATEPURTCPURTDRETURN(TC001, TC002, FORMID);
                //}
            }

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public void UPDATEPURMBPURMC(string TL001, string TL002, string FORMID)
        {
           

            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");


            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();

            //更新PURTL PURTM
            //新增PURMB PURMC
            queryString.AppendFormat(@"

                                    UPDATE [TK].dbo.PURTL
                                    SET TL006='Y',UDF02=@UDF02
                                    WHERE TL001=@TL001 AND TL002=@TL002

                                    UPDATE [TK].dbo.PURTM
                                    SET TM011='Y'
                                    WHERE TM001=@TL001 AND TM002=@TL002

                                      INSERT INTO  [TK].[dbo].[PURMB]
                                        (

                                        [COMPANY]
                                        ,[CREATOR]
                                        ,[USR_GROUP]
                                        ,[CREATE_DATE]
                                        ,[MODIFIER]
                                        ,[MODI_DATE]
                                        ,[FLAG]
                                        ,[CREATE_TIME]
                                        ,[MODI_TIME]
                                        ,[TRANS_TYPE]
                                        ,[TRANS_NAME]
                                        ,[sync_date]
                                        ,[sync_time]
                                        ,[sync_mark]
                                        ,[sync_count]
                                        ,[DataUser]
                                        ,[DataGroup]
                                        ,[MB001]
                                        ,[MB002]
                                        ,[MB003]
                                        ,[MB004]
                                        ,[MB005]
                                        ,[MB007]
                                        ,[MB008]
                                        ,[MB009]
                                        ,[MB010]
                                        ,[MB011]
                                        ,[MB012]
                                        ,[MB013]
                                        ,[MB014]
                                        ,[MB015]
                                        ,[MB016]
                                        ,[MB017]
                                        ,[MB018]
                                        ,[MB019]
                                        ,[MB020]
                                        ,[MB021]
                                        ,[MB022]
                                        ,[MB023]
                                        ,[MB024]
                                        ,[MB025]
                                        ,[MB026]
                                        ,[MB027]
                                        ,[MB028]
                                        ,[MB029]
                                        ,[MB030]
                                        ,[MB031]
                                        ,[MB032]
                                        ,[MB033]
                                        ,[MB034]
                                        ,[MB035]
                                        ,[MB036]
                                        ,[MB037]
                                        ,[MB038]
                                        ,[MB039]
                                        ,[MB040]
                                        ,[UDF01]
                                        ,[UDF02]
                                        ,[UDF03]
                                        ,[UDF04]
                                        ,[UDF05]
                                        ,[UDF06]
                                        ,[UDF07]
                                        ,[UDF08]
                                        ,[UDF09]
                                        ,[UDF10]
                                        )

                                        SELECT 
                                        PURTL.[COMPANY]  [COMPANY]
                                        ,PURTL.[CREATOR]   [CREATOR]
                                        ,PURTL.[USR_GROUP]   [USR_GROUP]
                                        ,PURTL.[CREATE_DATE]   [CREATE_DATE]
                                        ,PURTL.[MODIFIER]   [MODIFIER]
                                        ,PURTL.[MODI_DATE]   [MODI_DATE]
                                        ,PURTL.[FLAG]   [FLAG]
                                        ,PURTL.[CREATE_TIME]   [CREATE_TIME]
                                        ,PURTL.[MODI_TIME]   [MODI_TIME]
                                        ,'P004'   [TRANS_TYPE]
                                        ,'PURI03'   [TRANS_NAME]
                                        ,PURTL.[sync_date]   [sync_date]
                                        ,PURTL.[sync_time]   [sync_time]
                                        ,PURTL.[sync_mark]   [sync_mark]
                                        ,PURTL.[sync_count]   [sync_count]
                                        ,PURTL.[DataUser]   [DataUser]
                                        ,PURTL.[DataGroup]   [DataGroup]
                                        ,TM004 [MB001]
                                        ,TL004 [MB002]
                                        ,TL005 [MB003]
                                        ,TM009 [MB004]
                                        ,'' [MB005]
                                        ,TM007 [MB007]
                                        ,TL003 [MB008]
                                        ,'' [MB009]
                                        ,TM008 [MB010]
                                        ,TM010 [MB011]
                                        ,TM012+TL001+TL002 [MB012]
                                        ,TL008 [MB013]
                                        ,TM014 [MB014]
                                        ,TM015 [MB015]
                                        ,TM016 [MB016]
                                        ,'' [MB017]
                                        ,'' [MB018]
                                        ,'' [MB019]
                                        ,'' [MB020]
                                        ,'' [MB021]
                                        ,'1' [MB022]
                                        ,0 [MB023]
                                        ,0 [MB024]
                                        ,'' [MB025]
                                        ,0[MB026]
                                        ,'' [MB027]
                                        ,'' [MB028]
                                        ,'' [MB029]
                                        ,0 [MB030]
                                        ,0 [MB031]
                                        ,0 [MB032]
                                        ,0 [MB033]
                                        ,0 [MB034]
                                        ,'' [MB035]
                                        ,'' [MB036]
                                        ,'' [MB037]
                                        ,'' [MB038]
                                        ,'' [MB039]
                                        ,'' [MB040]
                                        ,'' [UDF01]
                                        ,'' [UDF02]
                                        ,'' [UDF03]
                                        ,'' [UDF04]
                                        ,'' [UDF05]
                                        ,0 [UDF06]
                                        ,0 [UDF07]
                                        ,0 [UDF08]
                                        ,0 [UDF09]
                                        ,0 [UDF10]
                                        FROM [TK].dbo.PURMA,[TK].dbo.PURTL,[TK].dbo.PURTM
                                        WHERE 1=1
                                        AND MA001=TL004
                                        AND TL001=TM001 AND TL002=TM002                                      
                                        AND LTRIM(RTRIM(TM004))+LTRIM(RTRIM(TL004))+LTRIM(RTRIM(TL005))+LTRIM(RTRIM(TM009))+LTRIM(RTRIM(TM014)) NOT IN (SELECT LTRIM(RTRIM(MB001))+LTRIM(RTRIM(MB002))+LTRIM(RTRIM(MB003))+LTRIM(RTRIM(MB004))+LTRIM(RTRIM(MB014)) FROM [TK].dbo.PURMB)
                                        
                                        AND TL001=@TL001 AND TL002=@TL002
                                   
                                        INSERT INTO [TK].[dbo].[PURMC]
                                        (
                                        [COMPANY]
                                        ,[CREATOR]
                                        ,[USR_GROUP]
                                        ,[CREATE_DATE]
                                        ,[MODIFIER]
                                        ,[MODI_DATE]
                                        ,[FLAG]
                                        ,[CREATE_TIME]
                                        ,[MODI_TIME]
                                        ,[TRANS_TYPE]
                                        ,[TRANS_NAME]
                                        ,[sync_date]
                                        ,[sync_time]
                                        ,[sync_mark]
                                        ,[sync_count]
                                        ,[DataUser]
                                        ,[DataGroup]
                                        ,[MC001]
                                        ,[MC002]
                                        ,[MC003]
                                        ,[MC004]
                                        ,[MC005]
                                        ,[MC006]
                                        ,[MC007]
                                        ,[MC008]
                                        ,[MC009]
                                        ,[MC010]
                                        ,[MC011]
                                        ,[MC012]
                                        ,[MC013]
                                        ,[MC014]
                                        ,[MC015]
                                        ,[MC016]
                                        ,[MC017]
                                        ,[MC018]
                                        ,[MC019]
                                        ,[MC020]
                                        ,[MC021]
                                        ,[MC022]
                                        ,[MC023]
                                        ,[MC024]
                                        ,[MC025]
                                        ,[MC026]
                                        ,[MC027]
                                        ,[MC028]
                                        ,[MC029]
                                        ,[MC030]
                                        ,[UDF01]
                                        ,[UDF02]
                                        ,[UDF03]
                                        ,[UDF04]
                                        ,[UDF05]
                                        ,[UDF06]
                                        ,[UDF07]
                                        ,[UDF08]
                                        ,[UDF09]
                                        ,[UDF10]
                                        )


                                        SELECT
                                        PURTL.[COMPANY]  [COMPANY]
                                        ,PURTL.[CREATOR]   [CREATOR]
                                        ,PURTL.[USR_GROUP]   [USR_GROUP]
                                        ,PURTL.[CREATE_DATE]   [CREATE_DATE]
                                        ,PURTL.[MODIFIER]   [MODIFIER]
                                        ,PURTL.[MODI_DATE]   [MODI_DATE]
                                        ,PURTL.[FLAG]   [FLAG]
                                        ,PURTL.[CREATE_TIME]   [CREATE_TIME]
                                        ,PURTL.[MODI_TIME]   [MODI_TIME]
                                        ,'P004'   [TRANS_TYPE]
                                        ,'PURI03'   [TRANS_NAME]
                                        ,PURTL.[sync_date]   [sync_date]
                                        ,PURTL.[sync_time]   [sync_time]
                                        ,PURTL.[sync_mark]   [sync_mark]
                                        ,PURTL.[sync_count]   [sync_count]
                                        ,PURTL.[DataUser]   [DataUser]
                                        ,PURTL.[DataGroup]   [DataGroup]
                                        ,TM004 [MC001]
                                        ,TL004 [MC002]
                                        ,TL005 [MC003]
                                        ,TM009 [MC004]
                                        ,TN007 [MC005]
                                        ,TN008 [MC006]
                                        ,TM012+TL001+TL002 [MC007]
                                        ,TM014 [MC008]
                                        ,'' [MC009]
                                        ,'' [MC010]
                                        ,'' [MC011]
                                        ,'' [MC012]
                                        ,'' [MC013]
                                        ,'1' [MC014]
                                        ,'' [MC015]
                                        ,0 [MC016]
                                        ,'' [MC017]
                                        ,'' [MC018]
                                        ,'' [MC019]
                                        ,0 [MC020]
                                        ,0 [MC021]
                                        ,0 [MC022]
                                        ,0 [MC023]
                                        ,0 [MC024]
                                        ,'' [MC025]
                                        ,'' [MC026]
                                        ,'' [MC027]
                                        ,'' [MC028]
                                        ,'' [MC029]
                                        ,'' [MC030]
                                        ,'' [UDF01]
                                        ,'' [UDF02]
                                        ,'' [UDF03]
                                        ,'' [UDF04]
                                        ,'' [UDF05]
                                        ,0 [UDF06]
                                        ,0 [UDF07]
                                        ,0 [UDF08]
                                        ,0 [UDF09]
                                        ,0 [UDF10]
                                        FROM [TK].dbo.PURMA,[TK].dbo.PURTL,[TK].dbo.PURTM
                                        LEFT JOIN [TK].dbo.PURTN ON TM001=TN001 AND TM002=TN002 AND TM003=TN003
                                        WHERE 1=1
                                        AND MA001=TL004
                                        AND TL001=TM001 AND TL002=TM002

                                        AND LTRIM(RTRIM(TM004))+LTRIM(RTRIM(TL004))+LTRIM(RTRIM(TL005))+LTRIM(RTRIM(TM009))+LTRIM(RTRIM(CONVERT(NVARCHAR,TN007)))+LTRIM(RTRIM(TM014)) NOT IN (SELECT LTRIM(RTRIM(MC001))+LTRIM(RTRIM(MC002))+LTRIM(RTRIM(MC003))+LTRIM(RTRIM(MC004))+LTRIM(RTRIM(CONVERT(NVARCHAR,MC005)))+LTRIM(RTRIM(MC008)) FROM [TK].dbo.PURMC)
                                        
                                        AND TL001=@TL001 AND TL002=@TL002
                                      
                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);

                    command.Parameters.Add("@TL001", SqlDbType.NVarChar).Value = TL001;
                    command.Parameters.Add("@TL002", SqlDbType.NVarChar).Value = TL002;
                  
                    command.Parameters.Add("@UDF02", SqlDbType.NVarChar).Value = FORMID;
           

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
