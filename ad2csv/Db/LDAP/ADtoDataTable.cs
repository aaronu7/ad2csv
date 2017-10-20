using System;
using System.Collections;
using System.Data;		// DataSets
using System.DirectoryServices;	//requires reference

using ad2csv.Db.LDAP;

namespace ad2csv
{
    public class ADtoDataTable
    {
        # region " AdGroups2DataTable "

        public void AdGroups2DataTable(DataTable oDt, string Domain, string SearchOU, string usr, string pwd) 
        {
            string ObjectClass = "group";
            string path = HelperLDAP.GetLdapPath(Domain, SearchOU);
			DirectoryEntry dePath = new DirectoryEntry(path, usr, pwd);

            SearchResultCollection oResults = HelperLDAP.ObjectFindSet(dePath, ObjectClass, "", "", "");
            foreach(SearchResult res in oResults) {
                DirectoryEntry oEntry = res.GetDirectoryEntry();

                DataRow oRow = oDt.NewRow();
                oDt.Rows.Add(oRow);

                string CN = "";
                string OU = "";
                string DC = "";
                HelperLDAP.GetPathParts(res.Path, "LDAP://" + Domain + "/", ref CN, ref OU, ref DC);

                oRow["path"] = res.Path;
                oRow["cn"] = CN;
                oRow["ou"] = OU;
                oRow["dc"] = DC;
                //oRow["membership"] = membership;
            }

        }

        # endregion

        # region " AdOrgUnits2DataTable "

        public void AdOrgUnits2DataTable(DataTable oDt, string Domain, string SearchOU, string usr, string pwd) 
        {
            string ObjectClass = "OrganizationalUnit";
            string path = HelperLDAP.GetLdapPath(Domain, SearchOU);
			DirectoryEntry dePath = new DirectoryEntry(path, usr, pwd);

            SearchResultCollection oResults = HelperLDAP.ObjectFindSet(dePath, ObjectClass, "", "", "");
            foreach(SearchResult res in oResults)
            {
                DirectoryEntry oEntry = res.GetDirectoryEntry();

                DataRow oRow = oDt.NewRow();
                oDt.Rows.Add(oRow);

                string CN = "";
                string OU = "";
                string DC = "";
                HelperLDAP.GetPathParts(res.Path, "LDAP://" + Domain + "/", ref CN, ref OU, ref DC);

                oRow["path"] = res.Path;
                oRow["cn"] = CN;
                oRow["ou"] = OU;
                oRow["dc"] = DC;
            }

        }

        # endregion

        # region " AdUsers2DataTable "

        public void AdUsers2DataTable(DataTable oDt, string Domain, string SearchOU, string usr, string pwd) 
        {
            string ObjectClass = "user";



            string path = HelperLDAP.GetLdapPath(Domain, SearchOU);
			DirectoryEntry dePath = new DirectoryEntry(path, usr, pwd);
            //DataTable oDt = BuildDt();

            SearchResultCollection oResults = HelperLDAP.ObjectFindSet(dePath, ObjectClass, "", "", "");
            foreach(SearchResult res in oResults)
            {
                DirectoryEntry oEntry = res.GetDirectoryEntry();
                //if(oEntry.NativeObject != null
                ArrayList memberSet = HelperLDAP.GroupsGetUserMemberships(oEntry);
                DataRow oRow = oDt.NewRow();
                oDt.Rows.Add(oRow);

                string membership = "";
                if (memberSet != null)
                {
                    foreach (DirectoryEntry obGpEntry in memberSet)
                    {
                        string gpCN = "";
                        string gpOU = "";
                        string gpDC = "";
                        HelperLDAP.GetPathParts(obGpEntry.Path, "LDAP://" + Domain + "/", ref gpCN, ref gpOU, ref gpDC);
                        if (membership == "")
                            membership = gpCN + "." + gpOU;
                        else
                            membership = membership + "," + gpCN + "." + gpOU;
                    }
                }

                string CN = "";
                string OU = "";
                string DC = "";
                HelperLDAP.GetPathParts(res.Path, "LDAP://" + Domain + "/", ref CN, ref OU, ref DC);

                oRow["path"] = res.Path;
                oRow["cn"] = CN;
                oRow["ou"] = OU;
                oRow["dc"] = DC;
                oRow["existLDAP"] = true;

                //------------------
                // Email
                string mail = "";
                if(oEntry.Properties.Contains("mail")) {
                    if (oEntry.Properties["mail"].Count > 0) {
                        System.DirectoryServices.PropertyValueCollection oPropSet = (System.DirectoryServices.PropertyValueCollection)oEntry.Properties["mail"];
                        //oProp = (object)oPropSet[0];
                        foreach(string val in oPropSet) {
                            if(val.Trim() != "") {
                                if(mail == "") { 
                                    mail = val;
                                } else {
                                    mail = mail + ";" + val;
                                }
                            }
                        }
                    }
                }
                oRow["mail"] = mail;

                //Object oMail = GetPropertyObject(oEntry, "mail");
                //if(oMail != null) {
                    //oRow["mail"] = GetPropertyAsString(oEntry, "mail");
                //    oRow["mail"] = (string)oMail;
                //}

                oRow["membership"] = membership;
                oRow["givenName"] = HelperLDAP.GetPropertyAsString(oEntry, "givenName");
                oRow["initials"] = HelperLDAP.GetPropertyAsString(oEntry, "initials");
                oRow["sn"] = HelperLDAP.GetPropertyAsString(oEntry, "sn");
                oRow["displayName"] = HelperLDAP.GetPropertyAsString(oEntry, "displayName");
                oRow["sAMAccountName"] = HelperLDAP.GetPropertyAsString(oEntry, "sAMAccountName");
                oRow["userPrincipalName"] = HelperLDAP.GetPropertyAsString(oEntry, "userPrincipalName");
                oRow["homeDrive"] = HelperLDAP.GetPropertyAsString(oEntry, "homeDrive");
                oRow["homeDirectory"] = HelperLDAP.GetPropertyAsString(oEntry, "homeDirectory");
                oRow["company"] = HelperLDAP.GetPropertyAsString(oEntry, "company");
                oRow["department"] = HelperLDAP.GetPropertyAsString(oEntry, "department");

                // The secondary design stuffed the Employee Number into extensionAttribute1
                string employeeID = HelperLDAP.GetPropertyAsString(oEntry, "employeeID");
                if(employeeID == "")
                    employeeID = HelperLDAP.GetPropertyAsString(oEntry, "extensionAttribute1");

                oRow["employeeID"] = employeeID;


                oRow["badPwdCount"] = HelperLDAP.GetPropertyAsInt32(oEntry, "badPwdCount");
                //long val = (long)oEntry.Properties["badPasswordTime"][0];
                //oRow["badPasswordTime"] = DateTime.FromFileTimeUtc(val);

                //long pwdLastSet = ConvertADSLargeIntegerToInt64(oEntry.Properties["pwdLastSet"].Value);
                //oRow["pwdLastSet"] = DateTime.FromFileTimeUtc(pwdLastSet);
                oRow["pwdLastSet"] = HelperLDAP.GetPropertyAsDateTime(oEntry, "pwdLastSet");

                //oRow["lastLogoff"] = (string)GetProperty(oEntry, "lastLogoff");
                //long lastLogoff = ConvertADSLargeIntegerToInt64(oEntry.Properties["lastLogoff"].Value);
                //oRow["lastLogoff"] = DateTime.FromFileTimeUtc(lastLogoff);
                oRow["lastLogoff"] = HelperLDAP.GetPropertyAsDateTime(oEntry, "lastLogoff");

                //oRow["lastLogon"] = (string)GetProperty(oEntry, "lastLogon");
                //long lastLogon = ConvertADSLargeIntegerToInt64(oEntry.Properties["lastLogon"].Value);
                //oRow["lastLogon"] = DateTime.FromFileTimeUtc(lastLogon);
                oRow["lastLogon"] = HelperLDAP.GetPropertyAsDateTime(oEntry, "lastLogon");

                //oRow["logonHours"] = (string)GetProperty(oEntry, "logonHours");   // tricky .... need to decode a byte array.
                oRow["logonCount"] = HelperLDAP.GetPropertyAsInt32(oEntry, "logonCount");

                // It's the number of ticks since Jan-02-1601.
                long accountExpiresTicks = HelperLDAP.ConvertADSLargeIntegerToInt64(oEntry.Properties["accountExpires"].Value);
                //Int64 accountExpiresTicks = (Int64)oEntry.Properties["accountExpires"].Value;
                //accountExpiresTicks = new DateTime(1601, 01, 02).AddTicks(accountExpiresTicks).Ticks;
                // 9223372032559808511  --- the value I'm getting
                // 9223372036854775807  --- found an online post saying this is the does not expire value
                //DateTime accountExpires = DateTime.MaxValue;
                //if (!accountExpiresTicks.Equals(Int64.MaxValue) && !accountExpiresTicks.Equals(9223372032559808511))
                //    accountExpires = DateTime.FromFileTime(accountExpiresTicks);
                //DateTime accountExpires = new DateTime(1601, 01, 02).AddTicks(accountExpiresTicks);
                //DateTime accountExpires = DateTime.FromFileTimeUtc(accountExpiresTicks);
                oRow["accountExpires"] = accountExpiresTicks;

                //oRow["lastLogonTimestamp"] = (string)GetProperty(oEntry, "lastLogonTimestamp");
                //long lastLogonTimestamp = ConvertADSLargeIntegerToInt64(oEntry.Properties["lastLogonTimestamp"].Value);
                //oRow["lastLogonTimestamp"] = DateTime.FromFileTimeUtc(lastLogonTimestamp);
                DateTime olastLogonTimestamp = HelperLDAP.GetPropertyAsDateTime(oEntry, "lastLogonTimestamp");
                DateTime olastLogonTimestampNew = new DateTime(olastLogonTimestamp.Year, olastLogonTimestamp.Month, olastLogonTimestamp.Day, olastLogonTimestamp.Hour, olastLogonTimestamp.Minute, olastLogonTimestamp.Second);
                oRow["lastLogonTimestamp"] = olastLogonTimestampNew;


            // lastLogonTimestamp has a strange timestamp format ..... should probably fix this in the code
            //DateTime oDt = DateTime.Parse("20/04/2015 5:56:13 PM");
            //DateTime oDt = DateTime.ParseExact("20/04/2015 5:56:13 PM", "dd/MM/yyyy HH:mm:ss tt",null);



                //if ((string)GetPropertyObject(oEntry, "employeeID") != "")
                //{
                    //foreach (string prop in oEntry.Properties.PropertyNames)
                    //{
                    //    System.Console.WriteLine(prop);
                    //}

                //    string a = "";
               // }

/*
                badPwdCount
                badPasswordTime
                lastLogoff
                lastLogon
                logonHours
                logonCount
                pwdLastSet
                accountExpires
                lastLogonTimestamp
*/

                /*
                objectClass
                cn
                description
                distinguishedName
                instanceType
                whenCreated
                whenChanged
                uSNCreated
                memberOf
                uSNChanged
                name
                objectGUID
                userAccountControl
                codePage
                countryCode
                primaryGroupID
                objectSid
                adminCount
                sAMAccountName
                sAMAccountType
                objectCategory
                isCriticalSystemObject
                dSCorePropagationData
                lastLogonTimestamp
                 */

                //Password, userPassword

                //byte[] pbytBytes = (byte[])aktBenutzer.Properties["userPassword"][0];
                byte[] pbytBytes = (byte[])HelperLDAP.GetPropertyObject(oEntry, "Password");
                if(pbytBytes != null)
                    oRow["password"] = pbytBytes.ToString();

                

                //oDt.Columns.Add(new DataColumn("password", Type.GetType("System.String")));
                //oDt.Columns.Add(new DataColumn("expire", Type.GetType("System.String")));       //true/false
                //oDt.Columns.Add(new DataColumn("enable", Type.GetType("System.String")));       //true/false
                //oDt.Columns.Add(new DataColumn("cantchange", Type.GetType("System.String")));   //true/false

                //oDt.Columns.Add(new DataColumn("membership", Type.GetType("System.String")));   //"students.groups.sd27,readers.groups.sd27,superusers.test.sd27"
            }
        }

        # endregion
    }
}
