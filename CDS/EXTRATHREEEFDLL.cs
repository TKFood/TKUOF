using Ede.Uof.EIP.Organization.Util;
using Ede.Uof.Utility.Data;
using Ede.Uof.WKF.CustomExternal;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TKUOF.CDS
{
    class EXTRATHREEEFDLL : ICallExternalDllSites
    {
        UserSet userSet1 = new UserSet();
        UserSet userSet2 = new UserSet();
        UserSet userSet3 = new UserSet();
        string singer1 = "";
        string singer2 = "";
        string singer3 = "";

        public void Finally()
        {
            //throw new NotImplementedException();
        }

        public string GetExternalDllSites(string formInfo)
        {
            XmlDocument formXmlDoc = new XmlDocument();
            DatabaseHelper DbQuery= new DatabaseHelper();
            UserUCO userUCO = new UserUCO();
            DataSet GroupName = new DataSet();
            string SpecialGroupName = "N";

            formXmlDoc.LoadXml(formInfo);
            string account = formXmlDoc.SelectSingleNode("/ExternalFlowSite/ApplicantInfo").Attributes["account"].Value;
            string userGuid = userUCO.GetGUID(account);
            EBUser ebUser = userUCO.GetEBUser(userGuid);
            string DEPNAME = ebUser.GroupName;

            GroupName = DbQuery.ExecuteDataSet("SELECT  [GroupName]   FROM [UOFTEST].[dbo].[GROUPNAMETOEF]");
            //判斷是不是指定部門，可由簽至副理就結束
            foreach (DataRow dr in GroupName.Tables[0].Rows)
            {
                if(DEPNAME.Equals(dr["GroupName"].ToString()))
                {
                    SpecialGroupName = "Y";
                    break;
                }
                else
                {
                    SpecialGroupName = "N";
                }
            }
               

            //找到申請者的第1層主管
            userSet1 =GetUserSuperior(userGuid);
            UserUCO userUCOsinger1 = new UserUCO();
            if (userSet1.Items.Count>0)
            {
                EBUser ebUsersinger1 = userUCOsinger1.GetEBUser(userSet1.Items[0].Key);
                singer1 = ebUsersinger1.Account;

                //主管的職級是理級以上就停止，不是就往上找
                if (ebUsersinger1.GetEmployeeDepartment(DepartmentOfUser.Major).JobTitle.Rank<=7)
                {
                    singer2 = null;
                }
                //如果指定部門的主管是副理以上就停止，不是就往上找
                else if (SpecialGroupName.Equals("Y") && ebUsersinger1.GetEmployeeDepartment(DepartmentOfUser.Major).JobTitle.Rank <= 9)
                {
                    singer2 = null;
                }
                else
                {
                    //找到申請者的第2層主管
                    userSet2 = GetUserSuperior(userSet1.Items[0].Key);
                    UserUCO userUCOsinger2 = new UserUCO();
                    if (userSet2.Items.Count > 0)
                    {
                        EBUser ebUsersinger2 = userUCOsinger2.GetEBUser(userSet2.Items[0].Key);
                        singer2 = ebUsersinger2.Account;

                        //主管的職級是理級以上就停止，不是就往上找
                        if (ebUsersinger2.GetEmployeeDepartment(DepartmentOfUser.Major).JobTitle.Rank <= 7)
                        {
                            singer3 = null;
                        }
                        //如果指定部門的主管是副理以上就停止，不是就往上找
                        else if (SpecialGroupName.Equals("Y") && ebUsersinger2.GetEmployeeDepartment(DepartmentOfUser.Major).JobTitle.Rank <=9)
                        {
                            singer3 = null;
                        }
                        else
                        {
                            //找到申請者的第3層主管
                            userSet3 = GetUserSuperior(userSet2.Items[0].Key);
                            UserUCO userUCOsinger3 = new UserUCO();
                            if (userSet3.Items.Count > 0)
                            {
                                EBUser ebUsersinger3 = userUCOsinger3.GetEBUser(userSet3.Items[0].Key);
                                singer3 = ebUsersinger3.Account;
                            }
                            else
                            {
                                singer3 = null;

                            }
                        }
                       
                    }
                    else
                    {
                        singer2 = null;

                    }
                }

               
            }
            else
            {
                singer1 = null;

            }



           
           

            

              


            //singer1 = "190052";
            //singer2 = "120002";
            //singer3 = "admin";

            //if (ebUser.HasJobFunction("Superior"))
            //{
            //    signer = "admin";
            //}
            //else
            //{
            //    signer = "Tony";
            //}

            XmlDocument xmlDoc = new XmlDocument();

            if (!string.IsNullOrEmpty(singer3))
            {                
                //<ReturnValue></ReturnValue>
                XmlElement returnValueElement = xmlDoc.CreateElement("ReturnValue");
                //<Flow></Flow>
                XmlElement flowElement = xmlDoc.CreateElement("Flow");
                //<Site></Site> 第1層
                XmlElement siteElement = xmlDoc.CreateElement("Site");
                //<Site order='' signType='' ></Site>
                siteElement.SetAttribute("order", "0");
                siteElement.SetAttribute("signType", "0");

                //<Signers></Signers>
                XmlElement signersElement = xmlDoc.CreateElement("Signers");
                //<Signer></Signer>
                XmlElement signerElement = xmlDoc.CreateElement("Signer");
                //<Signer account='' />
                signerElement.SetAttribute("account", singer1);
                //<Signers>
                //  <Signer account='' />
                //</Signers>
                signersElement.AppendChild(signerElement);
                //<Site order='' signType='' >
                //  <Signers>
                //      <Signer account='' />
                //  </Signers>
                //</Site>
                siteElement.AppendChild(signersElement);

                flowElement.AppendChild(siteElement);

                //<Site></Site> 第2層
                siteElement = xmlDoc.CreateElement("Site");
                //<Site order='' signType='' ></Site>
                siteElement.SetAttribute("order", "1");
                siteElement.SetAttribute("signType", "0");

                //<Signers></Signers>
                signersElement = xmlDoc.CreateElement("Signers");
                //<Signer></Signer>
                signerElement = xmlDoc.CreateElement("Signer");
                //<Signer account='' />
                signerElement.SetAttribute("account", singer2);
                //<Signers>
                //  <Signer account='' />
                //</Signers>
                signersElement.AppendChild(signerElement);
                //<Site order='' signType='' >
                //  <Signers>
                //      <Signer account='' />
                //  </Signers>
                //</Site>
                siteElement.AppendChild(signersElement);

                flowElement.AppendChild(siteElement);

                //<Site></Site> 第3層
                siteElement = xmlDoc.CreateElement("Site");
                //<Site order='' signType='' ></Site>
                siteElement.SetAttribute("order", "2");
                siteElement.SetAttribute("signType", "0");

                //<Signers></Signers>
                signersElement = xmlDoc.CreateElement("Signers");
                //<Signer></Signer>
                signerElement = xmlDoc.CreateElement("Signer");
                //<Signer account='' />
                signerElement.SetAttribute("account", singer3);
                //<Signers>
                //  <Signer account='' />
                //</Signers>
                signersElement.AppendChild(signerElement);
                //<Site order='' signType='' >
                //  <Signers>
                //      <Signer account='' />
                //  </Signers>
                //</Site>
                siteElement.AppendChild(signersElement);

                flowElement.AppendChild(siteElement);



                //<ReturnValue>
                //  <Flow>
                //    <Site order="0" signType="0">
                //      <Signers>
                //        <Signer account ="Tony"/>
                //      </Signers>
                //    </Site>

                //  </Flow>
                //</ReturnValue>


                returnValueElement.AppendChild(flowElement);
                xmlDoc.AppendChild(returnValueElement);
            }
            else if (!string.IsNullOrEmpty(singer2))
            {
                //<ReturnValue></ReturnValue>
                XmlElement returnValueElement = xmlDoc.CreateElement("ReturnValue");
                //<Flow></Flow>
                XmlElement flowElement = xmlDoc.CreateElement("Flow");
                //<Site></Site> 第1層
                XmlElement siteElement = xmlDoc.CreateElement("Site");
                //<Site order='' signType='' ></Site>
                siteElement.SetAttribute("order", "0");
                siteElement.SetAttribute("signType", "0");

                //<Signers></Signers>
                XmlElement signersElement = xmlDoc.CreateElement("Signers");
                //<Signer></Signer>
                XmlElement signerElement = xmlDoc.CreateElement("Signer");
                //<Signer account='' />
                signerElement.SetAttribute("account", singer1);
                //<Signers>
                //  <Signer account='' />
                //</Signers>
                signersElement.AppendChild(signerElement);
                //<Site order='' signType='' >
                //  <Signers>
                //      <Signer account='' />
                //  </Signers>
                //</Site>
                siteElement.AppendChild(signersElement);

                flowElement.AppendChild(siteElement);

                //<Site></Site> 第2層
                siteElement = xmlDoc.CreateElement("Site");
                //<Site order='' signType='' ></Site>
                siteElement.SetAttribute("order", "1");
                siteElement.SetAttribute("signType", "0");

                //<Signers></Signers>
                signersElement = xmlDoc.CreateElement("Signers");
                //<Signer></Signer>
                signerElement = xmlDoc.CreateElement("Signer");
                //<Signer account='' />
                signerElement.SetAttribute("account", singer2);
                //<Signers>
                //  <Signer account='' />
                //</Signers>
                signersElement.AppendChild(signerElement);
                //<Site order='' signType='' >
                //  <Signers>
                //      <Signer account='' />
                //  </Signers>
                //</Site>
                siteElement.AppendChild(signersElement);

                flowElement.AppendChild(siteElement);                

                returnValueElement.AppendChild(flowElement);
                xmlDoc.AppendChild(returnValueElement);
            }
            else if (!string.IsNullOrEmpty(singer1))
            {
                //<ReturnValue></ReturnValue>
                XmlElement returnValueElement = xmlDoc.CreateElement("ReturnValue");
                //<Flow></Flow>
                XmlElement flowElement = xmlDoc.CreateElement("Flow");
                //<Site></Site> 第1層
                XmlElement siteElement = xmlDoc.CreateElement("Site");
                //<Site order='' signType='' ></Site>
                siteElement.SetAttribute("order", "0");
                siteElement.SetAttribute("signType", "0");

                //<Signers></Signers>
                XmlElement signersElement = xmlDoc.CreateElement("Signers");
                //<Signer></Signer>
                XmlElement signerElement = xmlDoc.CreateElement("Signer");
                //<Signer account='' />
                signerElement.SetAttribute("account", singer1);
                //<Signers>
                //  <Signer account='' />
                //</Signers>
                signersElement.AppendChild(signerElement);
                //<Site order='' signType='' >
                //  <Signers>
                //      <Signer account='' />
                //  </Signers>
                //</Site>
                siteElement.AppendChild(signersElement);

                flowElement.AppendChild(siteElement);

                returnValueElement.AppendChild(flowElement);
                xmlDoc.AppendChild(returnValueElement);
            }

            return xmlDoc.OuterXml;
        }

        public void OnError(Exception errorException)
        {
            //throw new NotImplementedException();
        }

        // <summary>
        /// 取得員工直屬主管
        /// </summary>
        /// <param name="userGuid"></param>
        /// <returns></returns>
        public UserSet GetUserSuperior(string userGuid)
        {
            UserSet userSet = new UserSet();
            UserUCO userUCO = new UserUCO();
            EBUser ebUser = userUCO.GetEBUser(userGuid);
            BaseGroup baseGp = ebUser.GetEmployeeDepartment(DepartmentOfUser.Major).Department;

            //如果目前簽核人員不是部門主管，那就要找尋該人員目前部門的主管
            if (CheckIsDeptSuperior(ebUser.UserGUID, baseGp.GroupId) == false)
            {
                AddSuperiorToUserSet(baseGp.GroupId, userSet);
            }
            else
            {
                //如果目前的簽核人員是部門主管，則要找尋上一個部門的主管
                if (baseGp.ParnetGroup != null)
                {
                    AddSuperiorToUserSet(baseGp.ParnetGroup.GroupId, userSet);
                }
            }

            return userSet;

        }


        /// <summary>
        /// 檢查是否是否是部門主管
        /// </summary>
        /// <returns></returns>
        private bool CheckIsDeptSuperior(string userGUID, string groupId)
        {
            UserUCO userUco = new UserUCO();
            bool IsSuperior = false;
            EBUser ebUser = userUco.GetEBUser(userGUID);
            EmployeeJobFunctionCollection employeeJobFunctionCollection = ebUser.GetJobFunctionsOfDepartment(groupId);

            //判斷是否是簽核者，
            for (int i = 0; i < employeeJobFunctionCollection.Count; i++)
            {
                if (employeeJobFunctionCollection[i].FunctionId == "Superior")
                {
                    IsSuperior = true;
                    break;
                }               
            }
            
            return IsSuperior;
        }


        /// <summary>
        /// 把主管加到 UserSet 裡 , 2006/11/27 尋找部門主管，如果找不到就往上找，直到找到為止
        /// </summary>
        /// <param name="groupId"></param>
        /// <param name="userSet"></param>
        private void AddSuperiorToUserSet(string groupId, UserSet userSet)
        {
            EmployeeFindUCO employeeFindUCO = new EmployeeFindUCO();
            //取得查詢群組的主管
            UserSetEBUsersCollection userSetEBUsersCollection = employeeFindUCO.FindEmployeeByFunction(groupId, "Superior");

            if (userSetEBUsersCollection.Count > 0)
            {
                for (int i = 0; i < userSetEBUsersCollection.Count; i++)
                {
                    //把查到的主管 UserGuid 加到 userSet 裡
                    UserSetUser userSetUser = new UserSetUser();
                    userSetUser.USER_GUID = userSetEBUsersCollection[i].UserGUID;
                    userSet.Items.Add(userSetUser);
                }
            }
            else
            {
                //如果找不到直屬主管就往上一層層找上去，直到有為止
                BaseGroup baseGroup = new BaseGroup(groupId);
                if (baseGroup.ParnetGroup != null)
                {
                    AddSuperiorToUserSet(baseGroup.ParnetGroup.GroupId, userSet);
                }
            }
        }

    }
}
