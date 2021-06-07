using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using Ede.Uof.WKF.Utility;

namespace TKUOF.AUTONO
{
    class PURTABAUTONO : Ede.Uof.WKF.ExternalUtility.IFormAutoNumber
    {
        string TRACK_ID = "PURTA";
        public void Finally()
        {
            //throw new NotImplementedException();
        }

        public string GetFormNumber(string formId, string userGroupId, string formValueXML)
        {

            string connectionString = ConfigurationManager.ConnectionStrings["connectionstring"].ToString();
            Ede.Uof.Utility.Data.DatabaseHelper m_db = new Ede.Uof.Utility.Data.DatabaseHelper(connectionString);

            StringBuilder cmdTxt = new StringBuilder();
            cmdTxt.AppendFormat(@" 
                                SELECT TOP 1 [CURRENT_NO]
                                FROM [UOF].[dbo].[TK_WKF_AUTONO]
                                WHERE [TRACK_ID]='{0}' AND [AUTO_YEAR]='{1}' AND [AUTO_MONTH]='{2}' AND [AUTO_DAY]='{3}'
                                ", TRACK_ID, DateTime.Now.ToString("yyyy"), DateTime.Now.ToString("MM"), DateTime.Now.ToString("dd"));


            DataTable dt = new DataTable();

            dt.Load(m_db.ExecuteReader(cmdTxt.ToString()));

            if(dt.Rows.Count>=1)
            {
                string MAXNO= TRACK_ID + DateTime.Now.ToString("yyyy")+ DateTime.Now.ToString("MM") + DateTime.Now.ToString("dd")+ GETMAXNO(dt.Rows[0]["CURRENT_NO"].ToString());

                UPDATETK_WKF_AUTONO(MAXNO);

                return MAXNO;
            }
            else 
            {
                string MAXNO = TRACK_ID + DateTime.Now.ToString("yyyy") + DateTime.Now.ToString("MM") + DateTime.Now.ToString("dd") + "0001";

                INSERTTK_WKF_AUTONO(MAXNO);

                return MAXNO;
            }
            
            //throw new NotImplementedException();
        }

        public void OnError()
        {
            //throw new NotImplementedException();
        }

        public void OnExecption(Exception errorException)
        {
            //throw new NotImplementedException();
        }

        public string GETMAXNO(string CURRENT_NO)
        {
            if(!string.IsNullOrEmpty(CURRENT_NO))
            {
                string NO = CURRENT_NO.Substring(13,4);
                int MAXNO = Convert.ToInt32(NO) + 1;
                string RMAXNO= MAXNO.ToString().PadLeft(4, '0');

                return RMAXNO;

            }
            else
            {
                return "0001";
            }
        }

        public void UPDATETK_WKF_AUTONO(string MAXNO)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["connectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                        UPDATE [UOF].[dbo].[TK_WKF_AUTONO]
                                        SET [CURRENT_NO]=@MAXNO
                                        WHERE [TRACK_ID]='{0}' AND [AUTO_YEAR]='{1}' AND [AUTO_MONTH]='{2}' AND [AUTO_DAY]='{3}'
                                        ", TRACK_ID, DateTime.Now.ToString("yyyy"), DateTime.Now.ToString("MM"), DateTime.Now.ToString("dd"));

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@MAXNO", SqlDbType.NVarChar).Value = MAXNO;
          

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

        public void INSERTTK_WKF_AUTONO(string MAXNO)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["connectionstring"].ToString();

            StringBuilder queryString = new StringBuilder();
            queryString.AppendFormat(@"
                                        INSERT INTO  [UOF].[dbo].[TK_WKF_AUTONO]
                                        ([TRACK_ID],[AUTO_YEAR],[AUTO_MONTH],[AUTO_DAY],[CURRENT_NO])
                                        VALUES
                                        ('{0}','{1}','{2}','{3}',@MAXNO)
                                      ", TRACK_ID, DateTime.Now.ToString("yyyy"), DateTime.Now.ToString("MM"), DateTime.Now.ToString("dd"));

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@MAXNO", SqlDbType.NVarChar).Value = MAXNO;
                 

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
