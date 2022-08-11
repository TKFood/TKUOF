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

namespace TKUOF.TRIGGER.PURTEPURTF
{
    //採購變更的核準

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

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(applyTask.CurrentDocXML);
            TE001 = applyTask.Task.CurrentDocument.Fields["TE001"].FieldValue.ToString().Trim();
            TE002 = applyTask.Task.CurrentDocument.Fields["TE002"].FieldValue.ToString().Trim();
            TE003 = applyTask.Task.CurrentDocument.Fields["TE003"].FieldValue.ToString().Trim();
            FORMID = applyTask.FormNumber;

            ///核準
            if (applyTask.FormResult == Ede.Uof.WKF.Engine.ApplyResult.Adopt)
            {
                if (!string.IsNullOrEmpty(TE001) && !string.IsNullOrEmpty(TE002) && !string.IsNullOrEmpty(TE003))
                {
                    UPDATEPURTEPURTF(TE001, TE002, TE003, FORMID);
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

        public void UPDATEPURTEPURTF(string TE001, string TE002, string TE003, string FORMID)
        {
            string TC001 = TE001;
            string TC002 = TE002;
            string TD001 = TE001;
            string TD002 = TE002;          
            string TF001 = TE001;
            string TF002 = TE002;
            string TF003 = TE003;

            string COMPANY = "TK";
            string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            string MODI_TIME = DateTime.Now.ToString("HH:mm:dd");


            string connectionString = ConfigurationManager.ConnectionStrings["ERPconnectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                        --INSERT PURTD

                                        INSERT INTO [TK].dbo.PURTD
                                        (
                                        COMPANY,CREATOR,USR_GROUP,CREATE_DATE,FLAG,CREATE_TIME,MODI_TIME,TRANS_TYPE,TRANS_NAME,DataGroup
                                        ,TD001
                                        ,TD002
                                        ,TD003
                                        ,TD004
                                        ,TD005
                                        ,TD006
                                        ,TD007
                                        ,TD008
                                        ,TD009
                                        ,TD010
                                        ,TD011
                                        ,TD012
                                        ,TD014
                                        ,TD015
                                        ,TD016
                                        ,TD017
                                        ,TD018
                                        ,TD019
                                        ,TD020
                                        ,TD022
                                        ,TD025
                                        )

                                        SELECT 
                                        COMPANY,CREATOR,USR_GROUP,CREATE_DATE,FLAG,CREATE_TIME,MODI_TIME,TRANS_TYPE,TRANS_NAME,DataGroup
                                        ,TF001
                                        ,TF002
                                        ,TF104
                                        ,TF005
                                        ,TF006
                                        ,TF007
                                        ,TF008
                                        ,TF009
                                        ,TF010
                                        ,TF011
                                        ,TF012
                                        ,TF013
                                        ,TF030
                                        ,TF015
                                        ,'N'
                                        ,TF022
                                        ,'Y'
                                        ,TF018
                                        ,TF019
                                        ,TF020
                                        ,TF021
                                        FROM [TK].dbo.PURTF
                                        WHERE TF001=@TF001 AND TF002=@TF002 AND TF003=@TF003
                                        AND TF001+TF002+TF104 NOT IN (SELECT TD001+TD002+TD003  FROM [TK].dbo.PURTD WHERE TD001=@TD001 AND TD002=@TD002)

                                        --UPDATE PURTD

                                        UPDATE [TK].dbo.PURTD
                                        SET 
                                        TD004=TF005
                                        ,TD005=TF006
                                        ,TD006=TF007
                                        ,TD007=TF008
                                        ,TD008=TF009
                                        ,TD009=TF010
                                        ,TD010=TF011
                                        ,TD011=TF012
                                        ,TD012=TF013
                                        ,TD014=TF030
                                        ,TD015=TF015
                                        ,TD017=TF022
                                        ,TD019=TF018
                                        ,TD020=TF019
                                        ,TD022=TF020
                                        ,TD025=TF021
                                        FROM [TK].dbo.PURTF
                                        WHERE TD001=@TD001 AND TD002=@TD002 AND TD003=TF104
                                        AND TF001=@TF001 AND TF002=@TF002 AND TF003=@TF003


                                        --UPDATE PURTC

                                        UPDATE [TK].dbo.PURTC
                                        SET 
                                        TC004=TE005
                                        ,TC005=TE007
                                        ,TC006=TE008
                                        ,TC007=TE009
                                        ,TC008=TE010
                                        ,TC015=TE013
                                        ,TC016=TE014
                                        ,TC017=TE015
                                        ,TC018=TE018
                                        ,TC021=TE019
                                        ,TC022=TE020
                                        ,TC026=TE022
                                        ,TC027=TE023
                                        ,TC028=TE024
                                        ,TC009=TE027
                                        ,TC035=TE029
                                        ,TC011=TE037
                                        ,TC047=TE039
                                        ,TC048=TE040
                                        ,TC050=TE041
                                        ,TC036=TE043
                                        ,TC037=TE045
                                        ,TC038=TE046
                                        ,TC039=TE047
                                        ,TC040=TE048
                                        FROM [TK].dbo.PURTE
                                        WHERE TC001=@TC001 AND TC002=@TC002
                                        AND TE001=@TE001 AND TE002=@TE002 AND TE003=@TE003

                                        --更新PURTC的未稅、稅額、總金額、數量
                                        UPDATE [TK].dbo.PURTC
                                        SET TC019=(CASE WHEN TC018='1' THEN (SELECT ISNULL(ISNULL(ROUND(SUM(TD011)/(1+TC026),0),0),0) FROM [TK].dbo.PURTD WHERE TD001+TD002=TC001+TC002) 
                                                                            ELSE CASE WHEN TC018='2' THEN (SELECT ISNULL(SUM(TD011),0) FROM [TK].dbo.PURTD WHERE TD001+TD002=TC001+TC002) 
                                                                            ELSE CASE WHEN TC018='3' THEN (SELECT ISNULL(SUM(TD011),0) FROM [TK].dbo.PURTD WHERE TD001+TD002=TC001+TC002) 
                                                                            ELSE CASE WHEN TC018='4' THEN (SELECT ISNULL(SUM(TD011),0) FROM [TK].dbo.PURTD WHERE TD001+TD002=TC001+TC002) 
                                                                            ELSE CASE WHEN TC018='9' THEN (SELECT ISNULL(SUM(TD011),0) FROM [TK].dbo.PURTD WHERE TD001+TD002=TC001+TC002)  
                                                                            END
                                                                            END
                                                                            END 
                                                                            END
                                                                            END)
                                        ,TC020=(CASE WHEN TC018='1' THEN (SELECT (ISNULL(SUM(TD011),0)-ISNULL(ROUND(SUM(TD011)/(1+TC026),0),0)) FROM [TK].dbo.PURTD WHERE TD001+TD002=TC001+TC002) 
                                                                            ELSE CASE WHEN TC018='2' THEN (SELECT ISNULL(ROUND(SUM(TD011)*TC026,0),0) FROM [TK].dbo.PURTD WHERE TD001+TD002=TC001+TC002) 
                                                                            ELSE CASE WHEN TC018='3' THEN 0 
                                                                            ELSE CASE WHEN TC018='4' THEN 0
                                                                            ELSE CASE WHEN TC018='9' THEN 0 
                                                                            END
                                                                            END
                                                                            END 
                                                                            END
                                                                            END)
                                        ,TC023=(SELECT ISNULL(SUM(TD008),0) FROM [TK].dbo.PURTD WHERE TD001=TC001 AND TD002=TC002)
                                        WHERE TC001=@TC001 AND TC002=@TC002

                                        --如果變更單整理指定結案，原PURTC也指定結案

                                        UPDATE [TK].dbo.PURTD
                                        SET TD016='y'
                                        FROM [TK].dbo.PURTE
                                        WHERE TD001=@TD001 AND TD002=@TD002
                                        AND TE012='Y'                                    
                                        AND TE001=@TE001 AND TE002=@TE002 AND TE003=@TE003

                                        --如果變更單單身指定結案，原PURTD也指定結案
                                        UPDATE [TK].dbo.PURTD
                                        SET TD016='y'
                                        FROM [TK].dbo.PURTF
                                        WHERE   TD001=@TD001 AND TD002=@TD002
                                        AND TF001=TD001 AND TF002=TD002 AND TF104=TD003
                                        AND TF014='Y'                                       
                                        AND TF001=@TF001 AND TF002=@TF002 AND TF003=@TF003

                                        --更新PURTE
                                        UPDATE [TK].dbo.PURTE
                                        SET TE017='Y'
                                        ,UDF02=@UDF02
                                        WHERE TE001=@TE001 AND TE002=@TE002 AND TE003=@TE003

                                        --更新PURTF
                                        UPDATE [TK].dbo.PURTF
                                        SET TF016='Y'
                                        WHERE TF001=@TF001 AND TF002=@TF002 AND TF003=@TF003

                                        --更新PURTC
                                        UPDATE [TK].dbo.PURTC
                                        SET UDF03=@UDF03
                                        WHERE TC001=@TC001 AND TC002=@TC002
                                      
                                        ", FORMID);

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);

                    command.Parameters.Add("@TC001", SqlDbType.NVarChar).Value = TC001;
                    command.Parameters.Add("@TC002", SqlDbType.NVarChar).Value = TC002;
                    command.Parameters.Add("@TD001", SqlDbType.NVarChar).Value = TD001;
                    command.Parameters.Add("@TD002", SqlDbType.NVarChar).Value = TD002;
                    command.Parameters.Add("@TE001", SqlDbType.NVarChar).Value = TE001;
                    command.Parameters.Add("@TE002", SqlDbType.NVarChar).Value = TE002;
                    command.Parameters.Add("@TE003", SqlDbType.NVarChar).Value = TE003;
                    command.Parameters.Add("@TF001", SqlDbType.NVarChar).Value = TF001;
                    command.Parameters.Add("@TF002", SqlDbType.NVarChar).Value = TF002;
                    command.Parameters.Add("@TF003", SqlDbType.NVarChar).Value = TF003;
                    command.Parameters.Add("@UDF02", SqlDbType.NVarChar).Value = FORMID;
                    command.Parameters.Add("@UDF03", SqlDbType.NVarChar).Value = FORMID;

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
