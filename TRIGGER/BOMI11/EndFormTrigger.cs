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

namespace TKUOF.TRIGGER.BOMI11
{
    //ERP-BOM11.EBOM表 的 核準


    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {

        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string MJ001 = null;

            string FORMID = null;
            string MODIFIER = null;
            UserUCO userUCO = new UserUCO();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            MJ001 = applyTask.Task.CurrentDocument.Fields["MJ001"].FieldValue.ToString().Trim();
           

            FORMID = applyTask.FormNumber;
            //MODIFIER = applyTask.Task.Applicant.Account;

            //取得簽核人工號
            EBUser ebUser = userUCO.GetEBUser(Current.UserGUID);
            MODIFIER = ebUser.Account;

           
            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(MJ001) )
                {
                    UPDATE_BOMI11(MJ001, FORMID, MODIFIER);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATE_BOMI11(string MJ001, string FORMID, string MODIFIER)
        {           
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            string UDF01 = MODIFIER+"，已簽核:" +DateTime.Now.ToString("yyyyMMdd HH:mm:ss");
            string UDF02 = FORMID;

            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@" 
                                    INSERT  [TK].[dbo].[BOMMC]
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
                                    ,[UDF03]
                                    ,[UDF04]
                                    ,[UDF05]
                                        )
                                    SELECT 
                                    [BOMMJ].[COMPANY]
                                    ,[BOMMJ].[CREATOR]
                                    ,[BOMMJ].[USR_GROUP]
                                    ,[BOMMJ].[CREATE_DATE]
                                    ,[BOMMJ].[MODIFIER]
                                    ,[BOMMJ].[MODI_DATE]
                                    ,[BOMMJ].[FLAG]
                                    ,[BOMMJ].[CREATE_TIME]
                                    ,[BOMMJ].[MODI_TIME]
                                    ,[BOMMJ].[TRANS_TYPE]
                                    ,[BOMMJ].[TRANS_NAME]
                                    ,[BOMMJ].[sync_date]
                                    ,[BOMMJ].[sync_time]
                                    ,[BOMMJ].[sync_mark]
                                    ,[BOMMJ].[sync_count]
                                    ,[BOMMJ].[DataUser]
                                    ,[BOMMJ].[DataGroup]
                                    ,MJ001 [MC001]
                                    ,MB004 [MC002]
                                    ,MB072 [MC003]
                                    ,MJ004 [MC004]
                                    ,MJ005[MC005]
                                    ,'' [MC006]
                                    ,'' [MC007]
                                    ,'' [MC008]
                                    ,'0001' [MC009]
                                    ,MJ010 [MC010]
                                    ,0 [MC011]
                                    ,0 [MC012]
                                    ,'' [MC013]
                                    ,'' [MC014]
                                    ,'' [MC015]
                                    ,'N' [MC016]
                                    ,'' [MC017]
                                    ,'' [MC018]
                                    ,'N' [MC019]
                                    ,0 [MC020]
                                    ,0 [MC021]
                                    ,''[MC022]
                                    ,0 [MC023]
                                    ,MB002 [MC024]
                                    ,MB003 [MC025]
                                    ,'' [MC026]
                                    ,0 [MC027]
                                    ,[BOMMJ].[UDF03]
                                    ,[BOMMJ].[UDF04]
                                    ,[BOMMJ].[UDF05]
                                    
                                    
                                    FROM [TK].[dbo].[BOMMJ],[TK].[dbo].INVMB
                                    WHERE MJ001=MB001
                                    AND MJ001=@MJ001

                                    INSERT INTO [TK].[dbo].[BOMMD]
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
                                    ,[MD001]
                                    ,[MD002]
                                    ,[MD003]
                                    ,[MD004]
                                    ,[MD005]
                                    ,[MD006]
                                    ,[MD007]
                                    ,[MD008]
                                    ,[MD009]
                                    ,[MD010]
                                    ,[MD011]
                                    ,[MD012]
                                    ,[MD013]
                                    ,[MD014]
                                    ,[MD015]
                                    ,[MD016]
                                    ,[MD017]
                                    ,[MD018]
                                    ,[MD019]
                                    ,[MD020]
                                    ,[MD021]
                                    ,[MD022]
                                    ,[MD023]
                                    ,[MD024]
                                    ,[MD025]
                                    ,[MD026]
                                    ,[MD027]
                                    ,[MD028]
                                    ,[MD029]
                                    ,[MD030]
                                    ,[MD031]
                                    ,[MD032]
                                    ,[MD033]
                                    ,[MD034]
                                    ,[MD035]
                                    ,[MD036]
                                    ,[MD037]
                                    ,[MD038]
                                    )

                                    SELECT 
                                    [BOMMK].[COMPANY]
                                    ,[BOMMK].[CREATOR]
                                    ,[BOMMK].[USR_GROUP]
                                    ,[BOMMK].[CREATE_DATE]
                                    ,[BOMMK].[MODIFIER]
                                    ,[BOMMK].[MODI_DATE]
                                    ,[BOMMK].[FLAG]
                                    ,[BOMMK].[CREATE_TIME]
                                    ,[BOMMK].[MODI_TIME]
                                    ,[BOMMK].[TRANS_TYPE]
                                    ,[BOMMK].[TRANS_NAME]
                                    ,[BOMMK].[sync_date]
                                    ,[BOMMK].[sync_time]
                                    ,[BOMMK].[sync_mark]
                                    ,[BOMMK].[sync_count]
                                    ,[BOMMK].[DataUser]
                                    ,[BOMMK].[DataGroup]
                                    ,MK001 [MD001]
                                    ,MK002 [MD002]
                                    ,MK003 [MD003]
                                    ,MB004 [MD004]
                                    ,MB072 [MD005]
                                    ,MK006 [MD006]
                                    ,MK007 [MD007]
                                    ,MK008 [MD008]
                                    ,MK009 [MD009]
                                    ,'1' [MD010]
                                    ,MK011 [MD011]
                                    ,MK012 [MD012]
                                    ,MK013 [MD013]
                                    ,MK014 [MD014]
                                    ,MK015 [MD015]
                                    ,MK015 [MD016]
                                    ,MK017 [MD017]
                                    ,MK018 [MD018]
                                    ,MK020 [MD019]
                                    ,MK021 [MD020]
                                    ,MK022 [MD021]
                                    ,MK023 [MD022]
                                    ,MK024 [MD023]
                                    ,MK025 [MD024]
                                    ,MK026 [MD025]
                                    ,MK027 [MD026]
                                    ,MK028 [MD027]
                                    ,MK029 [MD028]
                                    ,'' [MD029]
                                    ,'' [MD030]
                                    ,'' [MD031]
                                    ,'N' [MD032]
                                    ,'' [MD033]
                                    ,0 [MD034]
                                    ,MB002 [MD035]
                                    ,MB003 [MD036]
                                    ,'' [MD037]
                                    ,0 [MD038]
                                    FROM [TK].[dbo].[BOMMK],[TK].[dbo].INVMB
                                    WHERE MK003=MB001
                                    AND MK001=@MK001

                                    UPDATE  [TK].dbo.BOMMJ
                                    SET
                                    UDF01=@UDF01
                                    ,UDF02=@UDF02
                                    WHERE MJ001=@MJ001 

                                    UPDATE  [TK].dbo.BOMMC
                                    SET
                                    UDF01=@UDF01
                                    ,UDF02=@UDF02
                                    WHERE MC001=@MC001

                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@MJ001", SqlDbType.NVarChar).Value = MJ001;
                    command.Parameters.Add("@MK001", SqlDbType.NVarChar).Value = MJ001;
                    command.Parameters.Add("@UDF01", SqlDbType.NVarChar).Value = UDF01;
                    command.Parameters.Add("@UDF02", SqlDbType.NVarChar).Value = UDF02;
                    command.Parameters.Add("@MC001", SqlDbType.NVarChar).Value = MJ001;

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
