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

                oRow["sAMAccountName"] = HelperLDAP.GetPropertyAsString(oEntry, "sAMAccountName");
                oRow["userPrincipalName"] = HelperLDAP.GetPropertyAsString(oEntry, "userPrincipalName");
                oRow["membership"] = membership;
                oRow["mail"] = HelperLDAP.GetPropertyAsStringSet(oEntry, "mail", ";");
                oRow["givenName"] = HelperLDAP.GetPropertyAsString(oEntry, "givenName");
                oRow["initials"] = HelperLDAP.GetPropertyAsString(oEntry, "initials");
                oRow["sn"] = HelperLDAP.GetPropertyAsString(oEntry, "sn");
                oRow["displayName"] = HelperLDAP.GetPropertyAsString(oEntry, "displayName");
                oRow["homeDrive"] = HelperLDAP.GetPropertyAsString(oEntry, "homeDrive");
                oRow["homeDirectory"] = HelperLDAP.GetPropertyAsString(oEntry, "homeDirectory");
                oRow["company"] = HelperLDAP.GetPropertyAsString(oEntry, "company");
                oRow["department"] = HelperLDAP.GetPropertyAsString(oEntry, "department");

                oRow["employeeID"] = HelperLDAP.GetPropertyAsString(oEntry, "employeeID");
                oRow["extensionAttribute1"] = HelperLDAP.GetPropertyAsString(oEntry, "extensionAttribute1");
                oRow["extensionAttribute2"] = HelperLDAP.GetPropertyAsString(oEntry, "extensionAttribute2");
                oRow["extensionAttribute3"] = HelperLDAP.GetPropertyAsString(oEntry, "extensionAttribute3");
            }
        }

        # endregion
    }
}
