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

namespace TKUOF.TRIGGER.COPTEF
{
    //訂單的核準

    class EndFormTrigger : ICallbackTriggerPlugin
    {
        public void Finally()
        {
            
        }

        public string GetFormResult(ApplyTask applyTask)
        {
            string TE001=null;
            string TE002 = null;
            string TE003 = null;
            string FORMID= null;
            string MODIFIER = null;
            string MOC = null;
            string PUR = null;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TE001 = applyTask.Task.CurrentDocument.Fields["TE001"].FieldValue.ToString().Trim();
            TE002 = applyTask.Task.CurrentDocument.Fields["TE002"].FieldValue.ToString().Trim();
            TE003 = applyTask.Task.CurrentDocument.Fields["TE003"].FieldValue.ToString().Trim();
            MOC = applyTask.Task.CurrentDocument.Fields["MOC"].FieldValue.ToString().Trim();
            PUR = applyTask.Task.CurrentDocument.Fields["PUR"].FieldValue.ToString().Trim();
            FORMID = applyTask.FormNumber;
            MODIFIER = applyTask.Task.Applicant.Account;

            ///核準 == Ede.Uof.WKF.Engine.ApplyResult.Adopt
            if (applyTask.SignResult== Ede.Uof.WKF.Engine.SignResult.Approve)
            {
                if (!string.IsNullOrEmpty(TE001) && !string.IsNullOrEmpty(TE002) && !string.IsNullOrEmpty(TE003))
                {
                    UPDATECOPTEF(TE001, TE002, TE003, FORMID, MODIFIER, MOC, PUR);
                }
            }
           

            return "";
        }

        public void OnError(Exception errorException)
        {
            
        }

        public void UPDATECOPTEF(string TE001, string TE002, string TE003, string FORMID, string MODIFIER,string MOC,string PUR)
        {
            string TE029 = "Y";
            string TE044 = "N";
            string TF019 = "Y";
            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");

            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                     --INSERT COPTD
                                    INSERT INTO [TK].dbo.COPTD
                                    (
                                    COMPANY,CREATOR,USR_GROUP,CREATE_DATE,FLAG,CREATE_TIME,MODI_TIME,TRANS_TYPE,TRANS_NAME,DataGroup
                                    ,TD001,TD002,TD003,TD004,TD005,TD006,TD007,TD008,TD010
                                    ,TD011,TD012,TD013,TD014,TD015,TD016,TD020
                                    ,TD021,TD023,TD024,TD026,TD027,TD028,TD029,TD030
                                    ,TD031,TD032,TD034,TD036,TD049,TD050,TD052
                                    ,TD088,TD089,TD090,TD091,TD092,TD093,TD094,TD095,TD096 ,TD097
                                    ) 

                                    SELECT COMPANY ,CREATOR ,USR_GROUP ,CREATE_DATE ,FLAG ,CREATE_TIME, MODI_TIME, TRANS_TYPE, TRANS_NAME,DataGroup
                                    ,TF001 TD001,TF002 TD002,TF104 TD003,TF005 TD004,TF006 TD005,TF007 TD006,TF008 TD007,TF009 TD008,TF010 TD010
                                    ,TF013 TD011,TF014 TD012,TF015 TD013,TF016 TD014,TF063 TD015,TF017 TD016,TF032 TD020
                                    ,'Y' TD021,TF012 TD023,TF020 TD026,TF021 TD027,TF022 TD028,TF064 TD029,TF023 TD029,TF024 TD030
                                    ,TF024 TD031,TF026 TD032,TF027 TD034,TF028 TD036,TF044 TD049,TF045 TD050,TF046 TD052
                                    ,TF183 TD088,TF184 TD089,TF185 TD090,TF186 TD091,TF187 TD092,TF188 TD093,TF189 TD094,TF194 TD095,TF195 TD096,TF137 TD097
                                    FROM [TK].dbo.COPTF
                                    WHERE TF001+TF002+TF104 NOT IN (SELECT TD001+TD002+TD003 FROM [TK].dbo.COPTD  WHERE TD001=@TD001 AND TD002=@TD002 )
                                    AND TF001=@TF001 AND TF002=@TF002 AND TF003=@TF003

                                    --更新COPTC
                                    UPDATE [TK].dbo.COPTC
                                    SET TC004=TE007,TC005=TE008,TC006=TE009,TC007=TE010,TC008=TE011
                                    ,TC009=TE012,TC010=TE013,TC011=TE014,TC012=TE015,TC013=TE016
                                    ,TC014=TE017,TC015=TE050,TC016=TE018,TC017=TE019,TC018=TE020
                                    ,TC019=TE021,TC020=TE022,TC021=TE023,TC022=TE024,TC023=TE025
                                    ,TC024=TE026,TC025=TE027,TC026=TE028,TC032=TE031,TC033=TE032
                                    ,TC034=TE033,TC035=TE034,TC036=TE035,TC037=TE036,TC038=TE037
                                    ,TC041=TE040,TC042=TE041,TC045=TE042,TC047=TE043,TC053=TE055
                                    ,TC054=TE047,TC055=TE048,TC056=TE049,TC057=TE183,TC058=TE184
                                    ,TC063=TE056,TC064=TE057,TC065=TE058,TC066=TE059,TC070=TE079 
                                    ,TC074=TE085,TC079=TE070,TC094=TE081,TC095=TE082,TC099=TE185
                                    ,TC104=TE084,TC105=TE086,TC113=TE087,TC114=TE088,TC115=TE196
                                    ,TC116=TE197
                                    ,FLAG=COPTC.FLAG+1
                                    FROM [TK].dbo.COPTE
                                    WHERE TC001=TE001 AND TC002=TE002
                                    AND TE001=@TE001 AND TE002=@TE002 AND TE003=@TE003

                                    --更新COPTD
                                    UPDATE [TK].dbo.COPTD
                                    SET TD004=TF005,TD005=TF006,TD006=TF007,TD007=TF008,TD008=TF009
                                    ,TD010=TF010,TD011=TF013,TD012=TF014,TD013=TF015,TD014=TF016
                                    ,TD015=TF063,TD016=TF017,TD020=TF032,TD023=TF012,TD024=TF020
                                    ,TD026=TF021,TD027=TF022,TD028=TF064,TD029=TF023,TD030=TF024
                                    ,TD031=TF025,TD032=TF026,TD034=TF027,TD036=TF028,TD049=TF044
                                    ,TD050=TF045,TD052=TF046,TD088=TF183,TD089=TF184,TD091=TF186
                                    ,TD092=TF187,TD093=TF188,TD094=TF189,TD095=TF194,TD096=TF195 
                                    ,TD097=TF137,TD090=TF185 
                                    ,FLAG=COPTD.FLAG+1
                                    FROM [TK].dbo.COPTF
                                    WHERE TD001+TD002+TD003=TF001+TF002+TF104
                                    AND TF001=@TF001 AND TF002=@TF002 AND TF003=@TF003

                                    UPDATE [TK].dbo.COPTE
                                    SET TE029=@TE029,TE044=@TE044, FLAG=FLAG+1,COMPANY=@COMPANY,MODIFIER=@MODIFIER ,MODI_DATE=@MODI_DATE, MODI_TIME=@MODI_TIME 
                                    AND TE001=@TE001 AND TE002=@TE002 AND TE003=@TE003

                                    UPDATE [TK].dbo.COPTF 
                                    SET TF019=@TF019, FLAG=FLAG+1,COMPANY=@COMPANY,MODIFIER=@MODIFIER ,MODI_DATE=@MODI_DATE, MODI_TIME=@MODI_TIME 
                                    AND TF001=@TF001 AND TF002=@TF002 AND TF003=@TF003

                                    --更新COPTC的未稅、稅額、總金額
                                    UPDATE [TK].dbo.COPTC
                                    SET TC029=(CASE WHEN TC016='1' THEN (SELECT ISNULL(ROUND(SUM(TD012)/(1+TC041),0),0) FROM [TK].dbo.COPTD WHERE TD001+TD002=TC001+TC002) 
                                    ELSE CASE WHEN TC016='2' THEN (SELECT ISNULL(SUM(TD012),0) FROM [TK].dbo.COPTD WHERE TD001+TD002=TC001+TC002) 
                                    ELSE CASE WHEN TC016='3' THEN (SELECT ISNULL(SUM(TD012),0) FROM [TK].dbo.COPTD WHERE TD001+TD002=TC001+TC002) 
                                    ELSE CASE WHEN TC016='4' THEN (SELECT ISNULL(SUM(TD012),0) FROM [TK].dbo.COPTD WHERE TD001+TD002=TC001+TC002) 
                                    ELSE CASE WHEN TC016='9' THEN (SELECT ISNULL(SUM(TD012),0) FROM [TK].dbo.COPTD WHERE TD001+TD002=TC001+TC002) 
                                    END
                                    END
                                    END 
                                    END
                                    END)
                                    ,TC030=(CASE WHEN TC016='1' THEN (SELECT (ISNULL(SUM(TD012),0)-ISNULL(ROUND(SUM(TD012)/(1+TC041),0),0)) FROM [TK].dbo.COPTD WHERE TD001+TD002=TC001+TC002) 
                                    ELSE CASE WHEN TC016='2' THEN (SELECT ISNULL(ROUND(SUM(TD012)*TC041,0),0) FROM [TK].dbo.COPTD WHERE TD001+TD002=TC001+TC002) 
                                    ELSE CASE WHEN TC016='3' THEN 0 
                                    ELSE CASE WHEN TC016='4' THEN 0
                                    ELSE CASE WHEN TC016='9' THEN 0 
                                    END
                                    END
                                    END 
                                    END
                                    END)
                                    ,TC031=(CASE WHEN TC016='1' THEN (SELECT ISNULL(SUM(TD012),0) FROM [TK].dbo.COPTD WHERE TD001+TD002=TC001+TC002) 
                                    ELSE CASE WHEN TC016='2' THEN (SELECT (ISNULL(SUM(TD012),0)+ISNULL(ROUND(SUM(TD012)*TC041,0),0)) FROM [TK].dbo.COPTD WHERE TD001+TD002=TC001+TC002) 
                                    ELSE CASE WHEN TC016='3' THEN (SELECT ISNULL(SUM(TD012),0) FROM [TK].dbo.COPTD WHERE TD001+TD002=TC001+TC002) 
                                    ELSE CASE WHEN TC016='4' THEN (SELECT ISNULL(SUM(TD012),0) FROM [TK].dbo.COPTD WHERE TD001+TD002=TC001+TC002) 
                                    ELSE CASE WHEN TC016='9' THEN (SELECT ISNULL(SUM(TD012),0) FROM [TK].dbo.COPTD WHERE TD001+TD002=TC001+TC002)  
                                    END
                                    END
                                    END 
                                    END
                                    END)
                                    WHERE TC001=@TE001 AND TC002=@TE002

                                    --更新COPTC總數量總數量、毛重(Kg)、材積(CUFT)
                                    UPDATE [TK].dbo.COPTC
                                    SET TC031=(SELECT ISNULL(SUM(TD008+TD024),0) FROM [TK].dbo.COPTD WHERE TD001+TD002=TC001+TC002)
                                    ,TC043=(SELECT ISNULL(SUM(TD030),0) FROM [TK].dbo.COPTD WHERE TD001+TD002=TC001+TC002)
                                    ,TC044=(SELECT ISNULL(SUM(TD031),0) FROM [TK].dbo.COPTD WHERE TD001+TD002=TC001+TC002)
                                    WHERE TC001=@TE001 AND TC002=@TE002

                                    UPDATE [TK].dbo.COPTD
                                    SET TD016='y'
                                    FROM [TK].dbo.COPTE
                                    WHERE TD001+TD002=TE001+TE002
                                    AND TE005='Y'
                                    WHERE TE001=@TE001 AND TE002=@TE002

                                        ");

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@TD001", SqlDbType.NVarChar).Value = TE001;
                    command.Parameters.Add("@TD002", SqlDbType.NVarChar).Value = TE002;
                    command.Parameters.Add("@TE001", SqlDbType.NVarChar).Value = TE001;
                    command.Parameters.Add("@TE002", SqlDbType.NVarChar).Value = TE002;
                    command.Parameters.Add("@TE003", SqlDbType.NVarChar).Value = TE003;
                    command.Parameters.Add("@TF001", SqlDbType.NVarChar).Value = TE001;
                    command.Parameters.Add("@TF002", SqlDbType.NVarChar).Value = TE002;
                    command.Parameters.Add("@TF003", SqlDbType.NVarChar).Value = TE003;
                    command.Parameters.Add("@TE029", SqlDbType.NVarChar).Value = TE029;
                    command.Parameters.Add("@TE044", SqlDbType.NVarChar).Value = TE044;
                    command.Parameters.Add("@TF019", SqlDbType.NVarChar).Value = TF019;
                    command.Parameters.Add("@FORMID", SqlDbType.NVarChar).Value = FORMID;                    
                    command.Parameters.Add("@COMPANY", SqlDbType.NVarChar).Value = COMPANY;
                    command.Parameters.Add("@MODIFIER", SqlDbType.NVarChar).Value = MODIFIER;
                    command.Parameters.Add("@MODI_DATE", SqlDbType.NVarChar).Value = MODI_DATE;
                    command.Parameters.Add("@MODI_TIME", SqlDbType.NVarChar).Value = MODI_TIME;
                    command.Parameters.Add("@MOC", SqlDbType.NVarChar).Value = MOC;
                    command.Parameters.Add("@PUR", SqlDbType.NVarChar).Value = PUR;

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
