using System;
using System.Collections;
using System.Data;		// DataSets
using System.DirectoryServices;	//requires reference

using ad2csv.Db.LDAP;

namespace ad2csv.Db.LDAP
{
    static public class Ad2DataTable
    {
        static public DataSet ExtractDataSet(string domain, string usr, string pwd, string ouLimit, int pageSize) {
            DataSet oDs = new DataSet();
            oDs.Tables.Add(ExtractDataTable_Groups(domain, usr, pwd, ouLimit, pageSize));
            oDs.Tables.Add(ExtractDataTable_Units(domain, usr, pwd, ouLimit, pageSize));
            oDs.Tables.Add(ExtractDataTable_Users(domain, usr, pwd, ouLimit, pageSize));
            return oDs;
        }

        # region " ExtractDataTable_Groups "

        static public DataTable ExtractDataTable_Groups(string Domain, string usr, string pwd, string ouLimit, int pageSize) 
        {
            string ObjectClass = "group";

            DataTable oDt = new DataTable("AD_Groups");
            oDt.Columns.Add(new DataColumn("path", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("dc", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("ou",     Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("cn", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("membership", Type.GetType("System.String")));

            string path = HelperLDAP.GetLdapPath(Domain, ouLimit);
			DirectoryEntry dePath = new DirectoryEntry(path, usr, pwd);

            SearchResultCollection oResults = HelperLDAP.ObjectFindSet(dePath, ObjectClass, "", pageSize);
            if(oResults != null) {
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
            return oDt;
        }

        # endregion

        # region " ExtractDataTable_Units "

        static public DataTable ExtractDataTable_Units(string Domain, string usr, string pwd, string ouLimit, int pageSize) 
        {
            string ObjectClass = "OrganizationalUnit";

            DataTable oDt = new DataTable("AD_Units");
            oDt.Columns.Add(new DataColumn("path", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("dc", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("ou",     Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("cn", Type.GetType("System.String")));

            string path = HelperLDAP.GetLdapPath(Domain, ouLimit);
			DirectoryEntry dePath = new DirectoryEntry(path, usr, pwd);

            SearchResultCollection oResults = HelperLDAP.ObjectFindSet(dePath, ObjectClass, "", pageSize);
            if(oResults != null) {
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
            return oDt;
        }

        # endregion

        # region " ExtractDataTable_Users "

        static public DataTable ExtractDataTable_Users(string Domain, string usr, string pwd, string ouLimit, int pageSize) 
        {
            string ObjectClass = "user";

            DataTable oDt = new DataTable("AD_Users");
            oDt.Columns.Add(new DataColumn("employeeID", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("cn", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("sAMAccountName", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("userPrincipalName", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("givenName", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("sn", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("displayName", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("mail", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("initials", Type.GetType("System.String")));            
            oDt.Columns.Add(new DataColumn("dc", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("ou",     Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("homeDrive", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("homeDirectory", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("company", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("department", Type.GetType("System.String")));       // SchoolName
            oDt.Columns.Add(new DataColumn("description", Type.GetType("System.String"))); 
            oDt.Columns.Add(new DataColumn("membership", Type.GetType("System.String")));   //"students.groups.sd27,readers.groups.sd27,superusers.test.sd27"
            oDt.Columns.Add(new DataColumn("path", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("extensionAttribute1", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("extensionAttribute2", Type.GetType("System.String")));
            oDt.Columns.Add(new DataColumn("extensionAttribute3", Type.GetType("System.String")));
            

            string path = HelperLDAP.GetLdapPath(Domain, ouLimit);
			DirectoryEntry dePath = new DirectoryEntry(path, usr, pwd);
            //DataTable oDt = BuildDt();

            SearchResultCollection oResults = HelperLDAP.ObjectFindSet(dePath, ObjectClass, "", pageSize);
            if(oResults != null) {
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
            return oDt;
        }

        # endregion
    }
}
