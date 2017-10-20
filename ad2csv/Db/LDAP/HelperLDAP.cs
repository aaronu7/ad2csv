
/****************************** DbHelper ******************************\
Module Name:  DbCsvPlus
Project:      This module is used in various ETL processes.
Copyright (c) Aaron Ulrich.


DbHelper contains various core functions to assist this process (ex. GetValueFromString)


This source is subject to the Apache License Version 2.0, January 2004
See http://www.apache.org/licenses/.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

using System;
using System.Collections;
using System.DirectoryServices;	//requires reference

namespace ad2csv.Db.LDAP
{
    static public class HelperLDAP {

        static public string GetObjectName(string ObjectClass, string CN) 
        {
            string ObjName = "";
            if(ObjectClass.ToLower()=="user" || ObjectClass.ToLower()=="group")
		        ObjName = "cn=" + CN;
            if(ObjectClass.ToLower() == "organizationalunit")
                ObjName = "ou=" + CN;
            return ObjName;
        }

        #region " GetUserMemberships "

        public static ArrayList GroupsGetUserMemberships(DirectoryEntry oUser) 
        {			
            ArrayList groups = new ArrayList();
            
            // Invoke Groups method.
			object obGroups = oUser.Invoke("Groups");
			foreach (object ob in (IEnumerable)obGroups)
			{
				DirectoryEntry obGpEntry = new DirectoryEntry(ob);
				groups.Add(obGpEntry);
			}

            return groups;
        }

        public static string GroupsGetUserMembershipsAsString(DirectoryEntry oUser, string Domain, bool IncludeOUPath) 
        {		
            ArrayList memberSet = GroupsGetUserMemberships(oUser);

            string membership = "";
            if (memberSet != null)
            {
                foreach (DirectoryEntry obGpEntry in memberSet)
                {
                    string gpCN = "";
                    string gpOU = "";
                    string gpDC = "";
                    GetPathParts(obGpEntry.Path, @"LDAP://" + Domain + @"/", ref gpCN, ref gpOU, ref gpDC);
                    if(IncludeOUPath) {
                        if (membership == "")
                            membership = gpCN + "." + gpOU;
                        else
                            membership = membership + "," + gpCN + "." + gpOU;
                    } else {
                        if (membership == "")
                            membership = gpCN;
                        else
                            membership = membership + "," + gpCN;
                    }
                     
                }
            }
            return membership;
        }

        #endregion

        #region " GetPropertyAs... "

        public static String GetPropertyAsStringSet(DirectoryEntry oEntry, string Prop, string div) {
            string res = "";
            if(oEntry.Properties.Contains(Prop)) {
                if (oEntry.Properties[Prop].Count > 0) {
                    System.DirectoryServices.PropertyValueCollection oPropSet = (System.DirectoryServices.PropertyValueCollection)oEntry.Properties["mail"];
                    foreach(string val in oPropSet) {
                        if(val.Trim() != "") {
                            if(res == "") { 
                                res = val;
                            } else {
                                res = res + div + val;
                            }
                        }
                    }
                }
            }
            return res;
        }

        public static String GetPropertyAsString(DirectoryEntry oEntry, string Prop) {
            object oProp = GetPropertyObject(oEntry, Prop);

            string val = "";
            if(oProp != null) {
                val = (String)oProp;
            }
            return val;
        }

        public static Int32 GetPropertyAsInt32(DirectoryEntry oEntry, string Prop) {
            object oProp = GetPropertyObject(oEntry, Prop);

            Int32 val = 0;
            if(oProp != null) {
                val = (Int32)oProp;
            }
            return val;
        }

        public static Int64 GetPropertyAsInt64(DirectoryEntry oEntry, string Prop) {
            object oProp = GetPropertyObject(oEntry, Prop);

            Int64 val = 0;
            if(oProp != null) {
                val = (Int64)oProp;
            }
            return val;
        }

        public static DateTime GetPropertyAsDateTime(DirectoryEntry oEntry, string Prop) {
            object oProp = GetPropertyObject(oEntry, Prop);

            DateTime val = DateTime.MinValue;
            if(oProp != null) {
                long dtAsLong = ConvertADSLargeIntegerToInt64(oEntry.Properties[Prop].Value);
                val = DateTime.FromFileTimeUtc(dtAsLong);
            }
            return val;
        }
        public static Int64 ConvertADSLargeIntegerToInt64(object adsLargeInteger)
        {
            if(adsLargeInteger == null) {
                return 0;
            } else {
              var highPart = (Int32)adsLargeInteger.GetType().InvokeMember("HighPart", System.Reflection.BindingFlags.GetProperty, null, adsLargeInteger, null);
              var lowPart  = (Int32)adsLargeInteger.GetType().InvokeMember("LowPart",  System.Reflection.BindingFlags.GetProperty, null, adsLargeInteger, null);
              return highPart * ((Int64)UInt32.MaxValue + 1) + lowPart;
            }
        }

        public static object GetPropertyObject(DirectoryEntry oEntry, string Prop)
        {
            object oProp = null;
            if(oEntry.Properties.Contains(Prop)) {
                if (oEntry.Properties[Prop].Count > 0) {
                    System.DirectoryServices.PropertyValueCollection oPropSet = (System.DirectoryServices.PropertyValueCollection)oEntry.Properties[Prop];
                    oProp = (object)oPropSet[0];
                }
            }
            return oProp;
        }

        #endregion

        #region " ObjectFindSet "

        static public SearchResultCollection ObjectFindSet(DirectoryEntry dePath, string ObjectClass, string ContextName, string ObjectCategory, string ANRFilter) 
        {
			DirectorySearcher deSearch = new DirectorySearcher();
            deSearch.PageSize = 100000;
			deSearch.SearchRoot = dePath;

            //dePath.Path

            //DirectorySearcher ds = new DirectorySearcher(
            //    entry,
            //    "(objectClass=organizationalUnit")
            //    null,
            //    SearchScope.OneLevel      // limits to searching only the immediate level
            //    );


            //ANR = Ambiguous Name Resolution - this allows you to search for a name without explicitly specifying which attribute to look it
            // ds.Filter = "(&(anr=" & txtSearch.Text & ")(objectCategory=person))"
            // mySearcher.Filter = ("(objectClass=computer)")

            //This filter here will select all the DISABLED accounts - that's those
            //that have the bit #2 in UserAccountControl set to true.
            //(userAccountControl:1.2.840.113556.1.4.803:=2)
            string ObjName = GetObjectName(ObjectClass, ContextName);
            string filter = "";
            if(ContextName == "")
                filter = "(&(objectClass=" + ObjectClass + "))";             
            else
                filter = "(&(objectClass=" + ObjectClass + ") (" + ObjName + "))";             

            deSearch.Filter = filter;
            SearchResultCollection results = null;
            try
            {
                // Can't quiet the error when the object is not found
                results = deSearch.FindAll();
            }
            catch { 
                System.Console.WriteLine("Exception caused by a failure to find the AD object.");
            }

            return results;
        }

        #endregion

        #region " GetPathParts "

        static public void GetPathParts(string Path, string domRoot, ref string CN, ref string OU, ref string DC)
        {
            // DomRoot = LDAP://root.sd27.bc.ca/
            // LDAP://root.sd27.bc.ca/CN=001,OU=numerics,OU=students,OU=sd27,DC=root,DC=sd27,DC=bc,DC=ca

            Path = Path.Replace(domRoot, "");
            string[] pathparts = Path.Split(',');
            foreach (string part in pathparts)
            {
                if (part.Trim() != "")
                {
                    string[] subparts = part.Split('=');
                    string sType = subparts[0].Trim();
                    string sName = subparts[1].Trim();

                    switch (sType.ToUpper())
                    {
                        case("DC"):
                            if (DC == "")
                                DC = sName;
                            else
                                DC = DC + "." + sName;
                            break;
                        case("OU"):
                            if (OU == "")
                                OU = sName;
                            else
                                OU = OU + "." + sName;
                            break;
                        case("CN"):
                            if (CN == "")
                                CN = sName;
                            else
                                CN = CN + "." + sName;
                            break;
                    }
                }
            }
        }

        #endregion

        #region " GetLdapPath "

        static public DirectoryEntry GetLdapPathEntry(string Domain, string OU, string usr, string pwd) 
        {
            string path = GetLdapPath(Domain, OU);
			DirectoryEntry entry = new DirectoryEntry(path, usr, pwd);
			return entry;
		}


        static public string GetLdapPath(string Domain, string OU) {
            return GetLdapPath(Domain, OU, "");
        }

        static public string GetLdapPath(string Domain, string OU, string UserOrGroup) {
            // Domain = domain.sd27.bc.ca
            // OU = students.sd27

            string LDAPDomain = GetLDAPDomain(Domain);
            string LDAPOU = GetLDAPOU(OU);
            string LDAPDomOU = "";
            string path = "";

            if(LDAPOU == "")
                LDAPDomOU = LDAPDomain;
            else
                LDAPDomOU = LDAPOU + "," + LDAPDomain;

            if(UserOrGroup == "")
                path ="LDAP://" + Domain + "/" + LDAPDomOU;
            else
                path ="LDAP://" + Domain + "/cn=" + UserOrGroup + "," + LDAPDomOU;


            return path;
        }

        static public string GetLDAPOU(string OU) 
        {
            // OU = students.sd27
            //  string LDAPOU: "OU=students,OU=sd27"
            string[] Parts = OU.Split('.');
            string Res = "";

            foreach(string part in Parts) {
                if(part != "") {
                    if(Res == "") {
                        Res = "OU=" + part;
                    } else {
                        Res = Res + ",OU=" + part;
                    }
                }
            }
            return Res;
        }

        static public string GetLDAPDomain(string Domain) 
        {
            // Domain = domain.sd27.bc.ca
            //      LDAP Domain = DC=domain,DC=sd27,DC=bc,DC=ca
            string[] LDAPDomainParts = Domain.Split('.');
            string LDAPDomain = "";
            foreach(string part in LDAPDomainParts) {
                if(LDAPDomain == "") {
                    LDAPDomain = "DC=" + part;
                } else {
                    LDAPDomain = LDAPDomain + ",DC=" + part;
                }
            }

            return LDAPDomain;
        }

        # endregion
    }
}
