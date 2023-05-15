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

namespace TKUOF.TRIGGER.PUR70
{
    //核準


    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {

        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TI001 = null;
            string TI002 = null;
            string FORMID = null;
            string MODIFIER = null;
            UserUCO userUCO = new UserUCO();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TI001 = applyTask.Task.CurrentDocument.Fields["TI001"].FieldValue.ToString().Trim();
            TI002 = applyTask.Task.CurrentDocument.Fields["TI002"].FieldValue.ToString().Trim();

            FORMID = applyTask.FormNumber;
            //MODIFIER = applyTask.Task.Applicant.Account;

            //取得簽核人工號
            EBUser ebUser = userUCO.GetEBUser(Current.UserGUID);
            MODIFIER = ebUser.Account;

           
            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(TI001) && !string.IsNullOrEmpty(TI002) )
                {
                    UPDATE_PUR70(TI001, TI002, FORMID, MODIFIER);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATE_PUR70(string TI001, string TI002, string FORMID, string MODIFIER)
        {           
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            string TI013 = "Y";
            string TI026 = MODIFIER;
            string TI032 = "N";
            string TJ020 = "Y";

            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"       
                                    UPDATE [test0923].dbo.PURTI
                                    SET
                                    TI013=@TI013 
                                    ,TI026=@TI026
                                    ,TI032=@TI032
                                    ,FLAG=FLAG+1
                                    ,MODIFIER=@MODIFIER 
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME 
                                    WHERE TI001=@TI001 AND TI002=TI002

                                    UPDATE [test0923].dbo.PURTJ
                                    SET
                                    TJ020=@TJ020
                                    ,FLAG=FLAG+1
                                    ,MODIFIER=@MODIFIER 
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME
                                    WHERE TJ001=@TJ001 AND TJ002=TJ002

                                    INSERT INTO [test0923].dbo.INVLA 
                                    (LA001 ,LA002 , LA003 ,LA004 ,LA005 ,LA006,LA007,LA008 ,LA009 ,LA010 , 
                                     LA011 ,LA012 ,LA013 ,LA014 ,LA015 ,LA016 ,LA017,LA018,LA019,LA020,LA021, 
                                     COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
                                     CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME ) 
                                    SELECT 
                                    TJ004 LA001 ,'' LA002 , '' LA003 ,TI003 LA004 ,-1 LA005 ,TJ001 LA006,TJ002 LA007,TJ003 LA008 ,TJ011 LA009 ,TI016 LA010 , 
                                    TJ009 LA011 ,TJ008 LA012 ,TJ010 LA013 ,1 LA014 , 'Y' LA015 ,TJ012 LA016 ,TJ010 LA017,0 LA018,0 LA019,0 LA020,0LA021, 
                                    PURTI.COMPANY ,PURTI.CREATOR ,PURTI.USR_GROUP ,PURTI.CREATE_DATE ,0 FLAG,  PURTI.CREATE_TIME, PURTI.MODI_TIME, PURTI.TRANS_TYPE, PURTI.TRANS_NAME 
                                    FROM [test0923].dbo.PURTI,[test0923].dbo.PURTJ
                                    WHERE TI001=TJ001 AND TI002=TJ002
                                    AND  TI001=@TI001 AND TI002=@TI002

                                    INSERT INTO [test0923].dbo.INVMF 
                                    (MF001 ,MF002 ,MF003 ,MF004 ,MF005 ,MF006,MF007,MF008 ,MF009 ,MF010 , 
                                     MF011 ,MF012 ,MF013,MF014 ,COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
                                     CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME )
                                    SELECT 
                                    TJ004 MF001 ,TJ012 MF002 ,TI003 MF003 ,TJ001 MF004 ,TJ002 MF005 ,TJ003 MF006,TJ011 MF007,-1 MF008 ,1 MF009 ,TJ009 MF010 , 
                                    '' MF011 ,'' MF012 ,TJ019 MF013,TJ024 MF014 
                                    ,PURTI.COMPANY ,PURTI.CREATOR ,PURTI.USR_GROUP ,PURTI.CREATE_DATE ,0 FLAG, 
                                    PURTI.CREATE_TIME, PURTI.MODI_TIME, PURTI.TRANS_TYPE, PURTI.TRANS_NAME
                                    FROM [test0923].dbo.PURTI,[test0923].dbo.PURTJ
                                    WHERE TI001=TJ001 AND TI002=TJ002
                                    AND  TI001=@TI001 AND TI002=@TI002

                                    UPDATE [test0923].dbo.PURTD 
                                    SET TD015=TD015-TJ009,TD016='N',TD031=TD031-TJ024 ,FLAG=FLAG+1
                                    FROM (
                                    SELECT 
                                    TJ004,SUM(TJ009) TJ009 ,SUM(TJ024) TJ024 ,TJ001,TJ002,TJ003,TJ016,TJ017,TJ018
                                    FROM [test0923].dbo.PURTI,[test0923].dbo.PURTJ
                                    WHERE TI001=TJ001 AND TI002=TJ002
                                    AND  TI001=@TI001 AND TI002=@TI002
                                    GROUP BY TJ004,TJ001,TJ002,TJ003,TJ016,TJ017,TJ018)
                                    AS TEMP
                                    WHERE TEMP.TJ016=TD001 AND  TEMP.TJ017=TD002 AND  TEMP.TJ018=TD003

                                    UPDATE [test0923].dbo.INVMB
                                    SET MB064=MB064-NUMS,MB065=MB065-TJ032
                                    FROM 
                                    (SELECT TJ004,SUM(TJ009) AS NUMS,SUM(TJ032) TJ032
                                    FROM [test0923].dbo.PURTI,[test0923].dbo.PURTJ
                                    WHERE TI001=TJ001 AND TI002=TJ002 
                                    AND  TI001=@TI001 AND TI002=@TI002
                                    GROUP BY TJ004) AS TEMP
                                    WHERE TEMP.TJ004=MB001

                                    UPDATE [test0923].dbo.INVMC
                                    SET MC007=MC007-NUMS,MC008=MC008-TJ032,MC012=TI003
                                    FROM 
                                    (SELECT TJ004,SUM(TJ009) AS NUMS,SUM(TJ032) TJ032,TJ011,TI003
                                    FROM [test0923].dbo.PURTI,[test0923].dbo.PURTJ
                                    WHERE TI001=TJ001 AND TI002=TJ002 
                                    AND  TI001=@TI001 AND TI002=@TI002
                                    GROUP BY TJ004,TJ011,TI003 ) AS TEMP
                                    WHERE TEMP.TJ004=MC001 AND TEMP.TJ011=MC002


                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TI001", SqlDbType.NVarChar).Value = TI001;
                    command.Parameters.Add("@TI002", SqlDbType.NVarChar).Value = TI002;

                    command.Parameters.Add("@TJ001", SqlDbType.NVarChar).Value = TI001;
                    command.Parameters.Add("@TJ002", SqlDbType.NVarChar).Value = TI002;

                    command.Parameters.Add("@TI013", SqlDbType.NVarChar).Value = TI013;
                    command.Parameters.Add("@TI026", SqlDbType.NVarChar).Value = TI026;
                    command.Parameters.Add("@TI032", SqlDbType.NVarChar).Value = TI032;
                    command.Parameters.Add("@TJ020", SqlDbType.NVarChar).Value = TJ020;
                 
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
