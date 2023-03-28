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

namespace TKUOF.TRIGGER.PUR60
{
    //ACRI02.結帳單 的核準


    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {

        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TG001 = null;
            string TG002 = null;
            string FORMID = null;
            string MODIFIER = null;
            UserUCO userUCO = new UserUCO();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TG001 = applyTask.Task.CurrentDocument.Fields["TG001"].FieldValue.ToString().Trim();
            TG002 = applyTask.Task.CurrentDocument.Fields["TG002"].FieldValue.ToString().Trim();

            FORMID = applyTask.FormNumber;
            //MODIFIER = applyTask.Task.Applicant.Account;

            //取得簽核人工號
            EBUser ebUser = userUCO.GetEBUser(Current.UserGUID);
            MODIFIER = ebUser.Account;

           
            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(TG001) && !string.IsNullOrEmpty(TG002) )
                {
                    UPDATE_PUR60(TG001, TG002, FORMID, MODIFIER);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATE_PUR60(string TG001, string TG002, string FORMID, string MODIFIER)
        {
           

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            string TG013 = "Y";
            string TG042 = "N";
            string TH030 = "Y";
            string TH038 = MODIFIER;
            string TH058 = "N";
            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"       
                                    UPDATE [test0923].dbo.PURTG
                                    SET
                                    TG013=@TG013
                                    ,TG042=@TG042 
                                    ,FLAG=FLAG+1
                                    ,MODIFIER=@MODIFIER
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME 
                                    WHERE TG001=@TG001 AND TG002=@TG002
 
                                    UPDATE [test0923].dbo.PURTH
                                    SET 
                                    TH030=@TH030 
                                    ,TH038=@TH038  
                                    ,TH058=@TH058
                                    ,FLAG=FLAG+1
                                    ,MODIFIER=@MODIFIER 
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME 
                                    WHERE TH001=@TH001 AND TH002=@TH002

                                    UPDATE [test0923].dbo.PURTD
                                    SET 
                                    TD015=TD015+NUMS
                                    ,TD016=( CASE WHEN TD008=TD015+NUMS THEN  'Y' END)
                                    ,TD031=0
                                    FROM 
                                    (
                                    SELECT TH001,TH002,TH003,TH011,TH012,TH013,(TH015-TH017) AS NUMS
                                    FROM [test0923].dbo.PURTG,[test0923].dbo.PURTH
                                    WHERE TG001=TH001 AND TG002=TH002 
                                    AND  TG001=@TG001 AND TG002=@TG002
                                    ) AS TEMP
                                    WHERE TD001=TH011 and TD002=TH012 and TD003=TH013

                                    INSERT INTO [test0923].dbo.INVLA
                                    (LA001 ,LA002 , LA003 ,LA004 ,LA005 ,LA006,LA007,LA008 ,LA009 ,LA010 , 
                                    LA011 ,LA012 ,LA013 ,LA014 ,LA015 ,LA016 ,LA017,LA018,LA019,LA020,LA021, 
                                    COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
                                    CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME ) 
                                    SELECT TH004 LA001 ,'' LA002 ,'' LA003 ,TG003 LA004 ,1 LA005 ,TH001 LA006,TH002 LA007,TH003 LA008 ,TH009 LA009 ,TH033 LA010 , 
                                    (TH015-TH017) LA011 ,TH018 LA012 ,TH047 LA013 ,'1' LA014 ,'Y' LA015 ,TH010 LA016 ,TH047 LA017,0 LA018,0 LA019,0 LA020,0 LA021, 
                                    PURTG.COMPANY ,PURTG.CREATOR CREATOR , PURTG.USR_GROUP  USR_GROUP ,  PURTG.CREATE_DATE CREATE_DATE ,0 FLAG, 
                                    PURTG.CREATE_TIME  CREATE_TIME, PURTG.MODI_TIME  MODI_TIME,'P004' TRANS_TYPE,'PURI09' TRANS_NAME
                                    FROM [test0923].dbo.PURTG,[test0923].dbo.PURTH
                                    WHERE TG001=TH001 AND TG002=TH002 
                                    AND  TG001=@TG001 AND TG002=@TG002

                                    INSERT INTO [test0923].dbo.INVME 
                                    (ME001,ME002,ME003,ME004,ME005,ME006,ME007,ME009,ME010,ME032
                                    ,FLAG,COMPANY,CREATOR,USR_GROUP,CREATE_DATE, CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME )
                                    SELECT TH004 ME001,TH010 ME002,TG003 ME003,'' ME004, TH001 ME005,TH002 ME006,'N' ME007,TH036 ME009,TH037 ME010,TH117 ME032
                                    ,0 FLAG,PURTG.COMPANY ,PURTG.CREATOR CREATOR , PURTG.USR_GROUP  USR_GROUP ,  PURTG.CREATE_DATE CREATE_DATE ,
                                    PURTG.CREATE_TIME  CREATE_TIME, PURTG.MODI_TIME  MODI_TIME,'P004' TRANS_TYPE,'PURI09' TRANS_NAME
                                    FROM [test0923].dbo.PURTG,[test0923].dbo.PURTH
                                    WHERE TG001=TH001 AND TG002=TH002 
                                    AND  TG001=@TG001 AND TG002=@TG002

                                    UPDATE [test0923].dbo.INVME 
                                    SET  
                                    ME003=TEMP.ME003,ME004=TEMP.ME004,ME005=TEMP.ME005
                                    ,ME006=TEMP.ME006,ME007=TEMP.ME007,ME008=TEMP.ME008,ME009=TEMP.ME009,ME010=TEMP.ME010, 
                                    ME032=TEMP.ME032
                                    FROM (SELECT TH004 ME001,TH010 ME002,TG003 ME003,'' ME004, TH001 ME005,TH002 ME006,'N' ME007,TH036 ME009,TH037 ME010,TH117 ME032,TH033 ME008
                                    ,0 FLAG,PURTG.COMPANY ,PURTG.CREATOR CREATOR , PURTG.USR_GROUP  USR_GROUP ,  PURTG.CREATE_DATE CREATE_DATE ,
                                    PURTG.CREATE_TIME  CREATE_TIME, PURTG.MODI_TIME  MODI_TIME,'P004' TRANS_TYPE,'PURI09' TRANS_NAME
                                    FROM [test0923].dbo.PURTG,[test0923].dbo.PURTH
                                    WHERE TG001=TH001 AND TG002=TH002 
                                    AND TG001=@TG001 AND TG002=@TG002
                                    ) AS TEMP
                                    WHERE INVME.ME001=TEMP.ME001 AND INVME.ME002=TEMP.ME002

                                    INSERT INTO [test0923].dbo.INVMF
                                     (MF001 ,MF002 ,MF003 ,MF004 ,MF005 ,MF006,MF007,MF008 ,MF009 ,MF010 , MF011 ,MF012 ,MF013,MF014 
                                    ,COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME ) 
                                    SELECT
                                    TH004 MF001 ,TH010 MF002 ,TG003 MF003 ,TH001 MF004 ,TH002 MF005 ,TH003 MF006,TH009 MF007,1 MF008 ,1 MF009 ,(TH015-TH017) MF010 ,'' MF011 ,'' MF012 ,'' MF013,0 F014 
                                    ,PURTG.COMPANY ,PURTG.CREATOR ,PURTG.USR_GROUP ,PURTG.CREATE_DATE ,0 FLAG, PURTG.CREATE_TIME, PURTG.MODI_TIME,'P004' TRANS_TYPE,'PURI09' TRANS_NAME
                                    FROM [test0923].dbo.PURTG,[test0923].dbo.PURTH
                                    WHERE TG001=TH001 AND TG002=TH002 
                                    AND  TG001=@TG001 AND TG002=@TG002
                                       


                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TG001", SqlDbType.NVarChar).Value = TG001;
                    command.Parameters.Add("@TG002", SqlDbType.NVarChar).Value = TG002;
                    command.Parameters.Add("@TG013", SqlDbType.NVarChar).Value = TG013;
                    command.Parameters.Add("@TG042", SqlDbType.NVarChar).Value = TG042;
                    command.Parameters.Add("@TH001", SqlDbType.NVarChar).Value = TG001;
                    command.Parameters.Add("@TH002", SqlDbType.NVarChar).Value = TG002;
                    command.Parameters.Add("@TH030", SqlDbType.NVarChar).Value = TH030;
                    command.Parameters.Add("@TH038", SqlDbType.NVarChar).Value = TH038;
                    command.Parameters.Add("@TH058", SqlDbType.NVarChar).Value = TH058;


                    command.Parameters.Add("@MODIFIER", SqlDbType.NVarChar).Value = MODIFIER;
                    command.Parameters.Add("@MODI_DATE", SqlDbType.NVarChar).Value = DateTime.Now.ToString("yyyyMMdd");
                    command.Parameters.Add("@MODI_TIME", SqlDbType.NVarChar).Value = DateTime.Now.ToString("HH:mm:ss");
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
