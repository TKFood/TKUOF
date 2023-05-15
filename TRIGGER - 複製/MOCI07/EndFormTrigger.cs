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

namespace TKUOF.TRIGGER.MOCI07
{
    //MOCTK MOCTL MOCI07.託外退貨單 的 核準


    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {

        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TK001 = null;
            string TK002 = null;
            string TK003 = null;
            string FORMID = null;
            string MODIFIER = null;
            UserUCO userUCO = new UserUCO();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TK001 = applyTask.Task.CurrentDocument.Fields["TK001"].FieldValue.ToString().Trim();
            TK002 = applyTask.Task.CurrentDocument.Fields["TK002"].FieldValue.ToString().Trim();
            TK003 = applyTask.Task.CurrentDocument.Fields["TK003"].FieldValue.ToString().Trim();

            FORMID = applyTask.FormNumber;
            //MODIFIER = applyTask.Task.Applicant.Account;

            //取得簽核人工號
            EBUser ebUser = userUCO.GetEBUser(Current.UserGUID);
            MODIFIER = ebUser.Account;

           
            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult == Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(TK001) && !string.IsNullOrEmpty(TK002) )
                {
                    UPDATE_MOCI07(TK001, TK002, TK003, FORMID, MODIFIER);
                }
            }


            return "";
        }

        public void OnError(Exception errorException)
        {

        }

        public void UPDATE_MOCI07(string TK001, string TK002, string TK003, string FORMID, string MODIFIER)
        {           
            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            string TK021 = "Y";
            string TK028 = MODIFIER;
            string TK035 = "N";


            string TL001 = TK001;
            string TL002 = TK002;
            string TL024 = "Y";


            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"   

                                    UPDATE [test0923].dbo.MOCTK
                                    SET
                                    TK021=@TK021 
                                    ,TK028=@TK028  
                                    ,TK035=@TK035 
                                    ,FLAG=FLAG+1
                                    ,MODIFIER=@MODIFIER 
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME
                                    WHERE TK001=@TK001 AND TK002=@TK002

                                    UPDATE [test0923].dbo.MOCTL
                                    SET
                                    TL024=@TL024 
                                    ,TL025='Y' 
                                    ,FLAG=FLAG+1
                                    ,MODIFIER=@MODIFIER 
                                    ,MODI_DATE=@MODI_DATE
                                    ,MODI_TIME=@MODI_TIME 
                                    WHERE TL001=@TL001 AND TL002=@TL002


                                    UPDATE [test0923].dbo.INVMB 
                                    SET 
                                    MB064=MB064+TUMS
                                    ,MB065=MB065+TMONEYS
                                    ,MB089=MB069+TPACKAGES 
                                    ,INVMB.FLAG=INVMB.FLAG+1
                                    FROM (
                                    SELECT TL004,SUM(-1*TL007) AS TUMS,(CASE WHEN MB064>0 AND MB065>0 THEN MB065/MB064*SUM(-1*TL007) ELSE 0 END) TMONEYS,SUM(-1*TL033) AS TPACKAGES
                                    FROM [test0923].dbo.MOCTL,[test0923].dbo.MOCTK,[test0923].dbo.INVMB,[test0923].dbo.CMSMQ
                                    WHERE TK001=TL001 AND TK002=TL002
                                    AND TL004=MB001
                                    AND MQ001=TK001
                                    AND TK001=@TK001 AND TK002=@TK002
                                    GROUP BY TL004,MQ010,MB064,MB065
                                    ) AS TEMP 
                                    WHERE TEMP.TL004=INVMB.MB001

                                    UPDATE [test0923].dbo.INVMC 
                                    SET
                                    MC007=MC007+TUMS
                                    ,MC008=MC008+TMONEYS
                                    ,MC014=MC014+TPACKAGES
                                    ,INVMC.FLAG=INVMC.FLAG+1
                                    FROM 
                                    (
                                    SELECT TL004,TL013,SUM(-1*TL007) AS TUMS,(CASE WHEN MB064>0 AND MB065>0 THEN MB065/MB064*SUM(-1*TL007) ELSE 0 END) TMONEYS,SUM(-1*TL033) AS TPACKAGES
                                    FROM [test0923].dbo.MOCTL,[test0923].dbo.MOCTK,[test0923].dbo.INVMB,[test0923].dbo.CMSMQ
                                    WHERE TK001=TL001 AND TK002=TL002
                                    AND TL004=MB001
                                    AND MQ001=TK001
                                    AND TK001=@TK001 AND TK002=@TK002
                                    GROUP BY TL004,TL013,MQ010,MB064,MB065
                                    )  AS TEMP
                                    WHERE TEMP.TL004=MC001 AND TEMP.TL013=MC002

                                    INSERT INTO [test0923].dbo.INVLA 
                                    (LA001 ,LA002 , LA003 ,LA004 ,LA005 ,LA006,LA007,LA008 ,LA009 ,LA010 , 
                                     LA011 ,LA012 ,LA013 ,LA014 ,LA015 ,LA016 ,LA017,LA018,LA019,LA020,LA021
                                     ,  COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
                                     CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME ) 
                                    SELECT TL004 LA001 ,'' LA002 ,'' LA003 ,TK003 LA004 ,'-1' LA005 ,TL001 LA006,TL002 LA007,TL003 LA008 ,TL013 LA009 ,MA002 LA010 
                                    ,TL007 LA011 ,(CASE WHEN MB064>0 AND MB065>0 THEN MB065/MB064 ELSE 0 END) LA012 ,(CASE WHEN MB064>0 AND MB065>0 THEN MB065/MB064*TL007 ELSE 0 END) LA013 ,'1' LA014 ,'Y' LA015 ,0 LA016 ,0 LA017,0 LA018,0 LA019,TL012 LA020,TL033 LA021
                                    ,MOCTK.COMPANY ,MOCTK.CREATOR ,MOCTK.USR_GROUP ,MOCTK.CREATE_DATE ,0 FLAG, 
                                    MOCTK.CREATE_TIME, MOCTK.MODI_TIME, 'P004' TRANS_TYPE,'MOCI07' TRANS_NAME 
                                    FROM [test0923].dbo.MOCTL,[test0923].dbo.MOCTK,[test0923].dbo.INVMB,[test0923].dbo.CMSMQ,[test0923].dbo.PURMA
                                    WHERE TK001=TL001 AND TK002=TL002
                                    AND TL004=MB001
                                    AND MQ001=TK001
                                    AND MA001=TK004
                                    AND TK001=@TK001 AND TK002=@TK002

                                    INSERT INTO [test0923].dbo.INVMF 
                                    (MF001 ,MF002 ,MF003 ,MF004 ,MF005 ,MF006,MF007,MF008 ,MF009 ,MF010 , 
                                     MF011 ,MF012 ,MF013,MF014 ,COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG, 
                                     CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME ) 
                                    SELECT TL004 MF001 ,TL014 MF002 ,TK003 MF003 ,TL001 MF004 ,TL002 MF005 ,TL003 MF006,TL013 MF007
                                    ,-1 MF008 ,MQ008 MF009 ,TL007 MF010 
                                    ,'' MF011 ,'' MF012 ,TK001+TK002 MF013,0 MF014 
                                    ,MOCTK.COMPANY ,MOCTK.CREATOR ,MOCTK.USR_GROUP ,MOCTK.CREATE_DATE ,0 FLAG
                                    ,MOCTK.CREATE_TIME, MOCTK.MODI_TIME, MOCTK.TRANS_TYPE, MOCTK.TRANS_NAME 
                                    FROM [test0923].dbo.MOCTL,[test0923].dbo.MOCTK,[test0923].dbo.INVMB,[test0923].dbo.CMSMQ
                                    WHERE TK001=TL001 AND TK002=TL002
                                    AND TL004=MB001
                                    AND MQ001=TK001
                                    AND TK001=@TK001 AND TK002=@TK002


                                    UPDATE [test0923].dbo.MOCTA 
                                    SET TA011='N'
                                    ,TA017=TA017-TNUMS
                                    ,TA046=TA046-TPACKAGES
                                    ,MOCTA.FLAG=MOCTA.FLAG+1
                                    FROM (
                                    SELECT TL015,TL016,SUM(TL007) AS TNUMS,SUM(TL033) AS TPACKAGES
                                    FROM [test0923].dbo.MOCTL,[test0923].dbo.MOCTK,[test0923].dbo.INVMB,[test0923].dbo.CMSMQ
                                    WHERE TK001=TL001 AND TK002=TL002
                                    AND TL004=MB001
                                    AND MQ001=TK001
                                    AND TK001=@TK001 AND TK002=@TK002
                                    GROUP BY TL015,TL016
                                    ) AS TEMP
                                    WHERE TEMP.TL015=MOCTA.TA001 AND TEMP.TL016=MOCTA.TA002






                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TK001", SqlDbType.NVarChar).Value = TK001;
                    command.Parameters.Add("@TK002", SqlDbType.NVarChar).Value = TK002;

                    command.Parameters.Add("@TL001", SqlDbType.NVarChar).Value = TK001;
                    command.Parameters.Add("@TL002", SqlDbType.NVarChar).Value = TK002;

                    command.Parameters.Add("@TK021", SqlDbType.NVarChar).Value = TK021;
                    command.Parameters.Add("@TK028", SqlDbType.NVarChar).Value = TK028;
                    command.Parameters.Add("@TK035", SqlDbType.NVarChar).Value = TK035;
                    command.Parameters.Add("@TL024", SqlDbType.NVarChar).Value = TL024;

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
