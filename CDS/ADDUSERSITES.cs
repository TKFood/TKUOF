using Ede.Uof.EIP.Organization.Util;
using Ede.Uof.Utility.Data;
using Ede.Uof.WKF.CustomExternal;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TKUOF.CDS
{
    class ADDUSERSITES : ICallExternalDllSites
    {
        Lib.WKF.ExternalDllSites sites = new Lib.WKF.ExternalDllSites();
        UserUCO UserUCOSuperior = new UserUCO();
        EBUser EBUserSuperior;
        UserSet userSet1 = new UserSet();

        Boolean FLAGGO = true;
        DataSet CompanyTopAccountDS = new DataSet();
        string account;
        string userGuid; 
        string CompanyTopAccount;
        string DEPNAME;
        string SpecialGroupName = "N";

        public void Finally()
        {
            
        }

        public string GetExternalDllSites(string formInfo)
        {
            XmlDocument xmlDoc = new XmlDocument();
            XmlDocument formXmlDoc = new XmlDocument();
            DatabaseHelper DbQueryCompanyTopAccount = new DatabaseHelper();
            DatabaseHelper DbQueryGroupName = new DatabaseHelper();
            UserUCO userUCO = new UserUCO();
            DataSet GroupName = new DataSet();

            formXmlDoc.LoadXml(formInfo);
            account = formXmlDoc.SelectSingleNode("/ExternalFlowSite/ApplicantInfo").Attributes["account"].Value;
            userGuid = userUCO.GetGUID(account);
            EBUser ebUser = userUCO.GetEBUser(userGuid);
            DEPNAME = ebUser.GroupName;

            //找出所有簽核人員，包含主管
            FINDALLSINGER(userGuid);

            return sites.ConvertToXML();
        }

        public void OnError(Exception errorException)
        {

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
