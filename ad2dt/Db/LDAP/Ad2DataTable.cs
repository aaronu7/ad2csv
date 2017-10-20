using System;
using System.Collections;
using System.Data;		// DataSets
using System.DirectoryServices;	//requires reference

using ad2csv.Db.LDAP;

namespace ad2csv.Db.LDAP
{
    /// <summary>
    /// This static class contains functions to extract DataTable listings of Active Directory objects and attributes.
    /// </summary>
    static public class Ad2DataTable
    {
        /// <summary>
        /// Use this function to extract a groups,orgunits and users from Active Directory and return them as a DataSet.
        /// </summary>
        /// <param name="Domain">The domain to connect to.</param>
        /// <param name="usr">The user account with rights to perform this operation.</param>
        /// <param name="pwd">The user accounts password.</param>
        /// <param name="ouLimit">Limit the ou scope. Format in the form of: Level3.Level2.Leve1</param>
        /// <param name="pageSize">Setting a pagesize will limit the total results returned from a query.</param>
        /// <returns>Returns a DataSet containing AD_Groups, AD_Units, AD_Users</returns>
        static public DataSet ExtractDataSet(string domain, string usr, string pwd, string ouLimit, int pageSize) {
            DataSet oDs = new DataSet();
            oDs.Tables.Add(ExtractDataTable_Groups(domain, usr, pwd, ouLimit, pageSize));
            oDs.Tables.Add(ExtractDataTable_Units(domain, usr, pwd, ouLimit, pageSize));
            oDs.Tables.Add(ExtractDataTable_Users(domain, usr, pwd, ouLimit, pageSize));
            return oDs;
        }

        # region " ExtractDataTable_Groups "

        /// <summary>
        /// Use this function to extract user groups from Active Directory and return them as a DataTable.
        /// </summary>
        /// <param name="Domain">The domain to connect to.</param>
        /// <param name="usr">The user account with rights to perform this operation.</param>
        /// <param name="pwd">The user accounts password.</param>
        /// <param name="ouLimit">Limit the ou scope. Format in the form of: Level3.Level2.Leve1</param>
        /// <param name="pageSize">Setting a pagesize will limit the total results returned from a query.</param>
        /// <returns>Returns a DataTable named AD_Groups</returns>
        static public DataTable ExtractDataTable_Groups(string Domain, string usr, string pwd, string ouLimit, int pageSize) 
        {
            string ObjectClass = "group";
            DataTable oDt = null;

            string path = HelperLdapGet.GetLdapPath(Domain, ouLimit);
			DirectoryEntry dePath = new DirectoryEntry(path, usr, pwd);

            SearchResultCollection oResults = HelperLdapGet.GetObjectSet(dePath, ObjectClass, "", pageSize);
            if(oResults != null) {

                oDt = new DataTable("AD_Groups");
                oDt.Columns.Add(new DataColumn("path", Type.GetType("System.String")));
                oDt.Columns.Add(new DataColumn("dc", Type.GetType("System.String")));
                oDt.Columns.Add(new DataColumn("ou",     Type.GetType("System.String")));
                oDt.Columns.Add(new DataColumn("cn", Type.GetType("System.String")));
                oDt.Columns.Add(new DataColumn("membership", Type.GetType("System.String")));

                foreach(SearchResult res in oResults) {
                    DirectoryEntry oEntry = res.GetDirectoryEntry();

                    DataRow oRow = oDt.NewRow();
                    oDt.Rows.Add(oRow);

                    string CN = "";
                    string OU = "";
                    string DC = "";
                    HelperLdapGet.GetLdapPathParts(res.Path, "LDAP://" + Domain + "/", ref CN, ref OU, ref DC);

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

        /// <summary>
        /// Use this function to extract the organizational units from Active Directory and return them as a DataTable.
        /// </summary>
        /// <param name="Domain">The domain to connect to.</param>
        /// <param name="usr">The user account with rights to perform this operation.</param>
        /// <param name="pwd">The user accounts password.</param>
        /// <param name="ouLimit">Limit the ou scope. Format in the form of: Level3.Level2.Leve1</param>
        /// <param name="pageSize">Setting a pagesize will limit the total results returned from a query.</param>
        /// <returns>Returns a DataTable named AD_Units</returns>
        static public DataTable ExtractDataTable_Units(string Domain, string usr, string pwd, string ouLimit, int pageSize) 
        {
            string ObjectClass = "OrganizationalUnit";
            DataTable oDt = null;

            string path = HelperLdapGet.GetLdapPath(Domain, ouLimit);
			DirectoryEntry dePath = new DirectoryEntry(path, usr, pwd);

            SearchResultCollection oResults = HelperLdapGet.GetObjectSet(dePath, ObjectClass, "", pageSize);
            if(oResults != null) {
                oDt = new DataTable("AD_Units");
                oDt.Columns.Add(new DataColumn("path", Type.GetType("System.String")));
                oDt.Columns.Add(new DataColumn("dc", Type.GetType("System.String")));
                oDt.Columns.Add(new DataColumn("ou",     Type.GetType("System.String")));
                oDt.Columns.Add(new DataColumn("cn", Type.GetType("System.String")));

                foreach(SearchResult res in oResults)
                {
                    DirectoryEntry oEntry = res.GetDirectoryEntry();

                    DataRow oRow = oDt.NewRow();
                    oDt.Rows.Add(oRow);

                    string CN = "";
                    string OU = "";
                    string DC = "";
                    HelperLdapGet.GetLdapPathParts(res.Path, "LDAP://" + Domain + "/", ref CN, ref OU, ref DC);

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

        /// <summary>
        /// Use this function to extract the users from Active Directory and return them as a DataTable.
        /// </summary>
        /// <param name="Domain">The domain to connect to.</param>
        /// <param name="usr">The user account with rights to perform this operation.</param>
        /// <param name="pwd">The user accounts password.</param>
        /// <param name="ouLimit">Limit the ou scope. Format in the form of: Level3.Level2.Leve1</param>
        /// <param name="pageSize">Setting a pagesize will limit the total results returned from a query.</param>
        /// <returns>Returns a DataTable named AD_Users</returns>
        static public DataTable ExtractDataTable_Users(string domain, string usr, string pwd, string ouLimit, int pageSize) 
        {
            string ObjectClass = "user";
            DataTable oDt = null;
            
            string path = HelperLdapGet.GetLdapPath(domain, ouLimit);
			DirectoryEntry dePath = new DirectoryEntry(path, usr, pwd);

            SearchResultCollection oResults = HelperLdapGet.GetObjectSet(dePath, ObjectClass, "", pageSize);
            if(oResults != null) {

                oDt = new DataTable("AD_Users");
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

                foreach(SearchResult res in oResults)
                {
                    DataRow oRow = oDt.NewRow();
                    oDt.Rows.Add(oRow);

                    DirectoryEntry oEntry = res.GetDirectoryEntry();

                    string cn = "";
                    string ou = "";
                    string dc = "";
                    HelperLdapGet.GetLdapPathParts(res.Path, "LDAP://" + domain + "/", ref cn, ref ou, ref dc);
                    string membership = HelperLdapGet.GroupsGetUserMembershipsAsString(oEntry, cn, ou, true, ",");

                    oRow["path"] = res.Path;
                    oRow["cn"] = cn;
                    oRow["ou"] = ou;
                    oRow["dc"] = dc;

                    oRow["sAMAccountName"] = HelperLdapGet.GetPropertyAsString(oEntry, "sAMAccountName");
                    oRow["userPrincipalName"] = HelperLdapGet.GetPropertyAsString(oEntry, "userPrincipalName");
                    oRow["membership"] = membership;
                    oRow["mail"] = HelperLdapGet.GetPropertyAsStringSet(oEntry, "mail", ";");
                    oRow["givenName"] = HelperLdapGet.GetPropertyAsString(oEntry, "givenName");
                    oRow["initials"] = HelperLdapGet.GetPropertyAsString(oEntry, "initials");
                    oRow["sn"] = HelperLdapGet.GetPropertyAsString(oEntry, "sn");
                    oRow["displayName"] = HelperLdapGet.GetPropertyAsString(oEntry, "displayName");
                    oRow["homeDrive"] = HelperLdapGet.GetPropertyAsString(oEntry, "homeDrive");
                    oRow["homeDirectory"] = HelperLdapGet.GetPropertyAsString(oEntry, "homeDirectory");
                    oRow["company"] = HelperLdapGet.GetPropertyAsString(oEntry, "company");
                    oRow["department"] = HelperLdapGet.GetPropertyAsString(oEntry, "department");

                    oRow["employeeID"] = HelperLdapGet.GetPropertyAsString(oEntry, "employeeID");
                    oRow["extensionAttribute1"] = HelperLdapGet.GetPropertyAsString(oEntry, "extensionAttribute1");
                    oRow["extensionAttribute2"] = HelperLdapGet.GetPropertyAsString(oEntry, "extensionAttribute2");
                    oRow["extensionAttribute3"] = HelperLdapGet.GetPropertyAsString(oEntry, "extensionAttribute3");
                }
            }
            return oDt;
        }

        # endregion
    }
}
