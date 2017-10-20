
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
using System.Collections.Generic;
using System.DirectoryServices;	//requires reference

namespace ad2csv.Db.LDAP
{
    /// <summary>
    /// This static helper class contains common LDAP helper and AD Get functions (Read-only)
    /// </summary>
    static public class HelperLdapGet {

        #region " GetObjectNameQuery "

        /// <summary>
        /// Format the ObjectName query portion of a directory object search.
        /// </summary>
        /// <param name="objectClass">The object class (ex: user, group, organizationalunit).</param>
        /// <param name="cn">The CN of the object.</param>
        /// <returns></returns>
        static public string GetObjectNameQuery(string objectClass, string cn) {
            string ObjName = "";
            if(objectClass.ToLower()=="user" || objectClass.ToLower()=="group")
		        ObjName = "cn=" + cn;
            if(objectClass.ToLower() == "organizationalunit")
                ObjName = "ou=" + cn;
            return ObjName;
        }

        #endregion

        #region " GetUserMemberships "

        /// <summary>
        /// Get a list of a users memberships.
        /// </summary>
        /// <param name="oUser">The users directory object.</param>
        /// <returns>A list of DirectoryEntry objects.</returns>
        public static List<DirectoryEntry> GroupsGetUserMemberships(DirectoryEntry oUser) {			
            List<DirectoryEntry> groups = new List<DirectoryEntry>();
            
            // Invoke Groups method.
			object obGroups = oUser.Invoke("Groups");
			foreach (object ob in (IEnumerable)obGroups) {
				DirectoryEntry obGpEntry = new DirectoryEntry(ob);
				groups.Add(obGpEntry);
			}

            return groups;
        }

        /// <summary>
        /// Compose a string delimited list of a users memberships.
        /// </summary>
        /// <param name="oUser"></param>
        /// <param name="Domain"></param>
        /// <param name="IncludeOUPath"></param>
        /// <returns></returns>
        public static string GroupsGetUserMembershipsAsString(DirectoryEntry oUser, string cn, string ou, bool includeOUPath, string delim) 
        {		
            List<DirectoryEntry> memberSet = GroupsGetUserMemberships(oUser);

            string membership = "";
            if (memberSet != null) {
                foreach (DirectoryEntry obGpEntry in memberSet) {
                    if(includeOUPath) {
                        if (membership == "")
                            membership = cn + "." + ou;
                        else
                            membership = membership + "," + cn + "." + ou;
                    } else {
                        if (membership == "")
                            membership = cn;
                        else
                            membership = membership + delim + cn;
                    }                     
                }
            }
            return membership;
        }

        #endregion

        #region " GetPropertyAs... "

        /// <summary>
        /// Get a PropertyValueCollection as a string delimited set.
        /// </summary>
        /// <param name="oEntry">The directory object entry.</param>
        /// <param name="Prop">The property to get.</param>
        /// <param name="div">The delimiter to seperate property values by.</param>
        /// <returns></returns>
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

        /// <summary>
        /// Get a PropertyValue as a string.
        /// </summary>
        /// <param name="oEntry">The directory object entry.</param>
        /// <param name="Prop">The property to get.</param>
        /// <returns></returns>
        public static String GetPropertyAsString(DirectoryEntry oEntry, string Prop) {
            object oProp = GetPropertyObject(oEntry, Prop);

            string val = "";
            if(oProp != null) {
                val = (String)oProp;
            }
            return val;
        }

        /// <summary>
        /// Get a PropertyValue as a Int32.
        /// </summary>
        /// <param name="oEntry">The directory object entry.</param>
        /// <param name="Prop">The property to get.</param>
        /// <returns></returns>
        public static Int32 GetPropertyAsInt32(DirectoryEntry oEntry, string Prop) {
            object oProp = GetPropertyObject(oEntry, Prop);

            Int32 val = 0;
            if(oProp != null) {
                val = (Int32)oProp;
            }
            return val;
        }

        /// <summary>
        /// Get a PropertyValue as a Int64.
        /// </summary>
        /// <param name="oEntry">The directory object entry.</param>
        /// <param name="Prop">The property to get.</param>
        /// <returns></returns>
        public static Int64 GetPropertyAsInt64(DirectoryEntry oEntry, string Prop) {
            object oProp = GetPropertyObject(oEntry, Prop);

            Int64 val = 0;
            if(oProp != null) {
                val = (Int64)oProp;
            }
            return val;
        }

        /// <summary>
        /// Get a PropertyValue as a DateTime.
        /// </summary>
        /// <param name="oEntry">The directory object entry.</param>
        /// <param name="Prop">The property to get.</param>
        /// <returns></returns>
        public static DateTime GetPropertyAsDateTime(DirectoryEntry oEntry, string Prop) {
            object oProp = GetPropertyObject(oEntry, Prop);

            DateTime val = DateTime.MinValue;
            if(oProp != null) {
                long dtAsLong = ConvertADSLargeIntegerToInt64(oEntry.Properties[Prop].Value);
                val = DateTime.FromFileTimeUtc(dtAsLong);
            }
            return val;
        }

        /// <summary>
        /// Covert and ADS Large Integer into an Int64
        /// </summary>
        /// <param name="adsLargeInteger">The large integer object.</param>
        /// <returns></returns>
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

        /// <summary>
        /// Get a PropertyValue as a Object.
        /// </summary>
        /// <param name="oEntry">The directory object entry.</param>
        /// <param name="Prop">The property to get.</param>
        /// <returns></returns>
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

        #region " GetObjectSet "

        /// <summary>
        /// This function will simply return the count of objects matching the scope criteria.
        /// </summary>
        /// <param name="domain">The domain</param>
        /// <param name="ouLimit">The ou limiter.</param>
        /// <param name="usr">Username with proper credentials.</param>
        /// <param name="pwd">Username password.</param>
        /// <param name="objectClass">The object class (ex: user, group, organizationalunit).</param>
        /// <param name="contextName">The CN for the object.</param>
        /// <param name="pageSize">A pagesize limiter.</param>
        /// <returns></returns>
        static public int GetObjectSetCount(string domain, string ouLimit, string usr, string pwd, string objectClass, string contextName, int pageSize)  {

            string path = HelperLdapGet.GetLdapPath(domain, ouLimit);
			DirectoryEntry dePath = new DirectoryEntry(path, usr, pwd);

            SearchResultCollection results = GetObjectSet(dePath, objectClass, contextName, pageSize);
            if(results==null)
                return 0;
            else
                return results.Count;
        }

        /// <summary>
        /// This function will return a collection of DirectoryEntry objects that meet the scope criteria. 
        /// </summary>
        /// <param name="oDomainPath">The search root with domain and ou limiters applied.</param>
        /// <param name="ObjectClass">The object class (ex: user, group, organizationalunit).</param>
        /// <param name="ContextName">The CN for the object.</param>
        /// <param name="pageSize">A pagesize limiter.</param>
        /// <returns></returns>
        static public SearchResultCollection GetObjectSet(DirectoryEntry oDomainPath, string objectClass, string contextName, int pageSize) 
        {
			DirectorySearcher deSearch = new DirectorySearcher();
            deSearch.PageSize = pageSize;
			deSearch.SearchRoot = oDomainPath;
            
            string ObjName = GetObjectNameQuery(objectClass, contextName);
            string filter = "";
            if(contextName == "")
                filter = "(&(objectClass=" + objectClass + "))";             
            else
                filter = "(&(objectClass=" + objectClass + ") (" + ObjName + "))";             

            deSearch.Filter = filter;
            SearchResultCollection results = null;
            try {
                // Can't quiet the error when the object is not found
                results = deSearch.FindAll();
            } catch { 
                System.Console.WriteLine("Exception caused by a failure to find the AD object.");
            }

            return results;
        }

        #endregion

        #region " GetLdapPathParts "

        /// <summary>
        /// Takes an LDAP path and populates the references to cn, ou and dc.
        /// </summary>
        /// <param name="ldapPath">The LDAP path.</param>
        /// <param name="domRoot">The root of the LDAP path (ex. LDAP://domain.bc.ca/) </param>
        /// <param name="cn">The cn value to populate.</param>
        /// <param name="ou">The ou value to populate.</param>
        /// <param name="dc">The dc value to populate.</param>
        static public void GetLdapPathParts(string ldapPath, string domRoot, ref string cn, ref string ou, ref string dc)
        {
            // DomRoot = LDAP://root.sd27.bc.ca/
            // LDAP://root.sd27.bc.ca/CN=001,OU=numerics,OU=students,OU=sd27,DC=root,DC=sd27,DC=bc,DC=ca

            ldapPath = ldapPath.Replace(domRoot, "");
            string[] pathparts = ldapPath.Split(',');
            foreach (string part in pathparts) {
                if (part.Trim() != "") {
                    string[] subparts = part.Split('=');
                    string sType = subparts[0].Trim();
                    string sName = subparts[1].Trim();

                    switch (sType.ToUpper()) {
                        case("DC"):
                            if (dc == "")
                                dc = sName;
                            else
                                dc = dc + "." + sName;
                            break;
                        case("OU"):
                            if (ou == "")
                                ou = sName;
                            else
                                ou = ou + "." + sName;
                            break;
                        case("CN"):
                            if (cn == "")
                                cn = sName;
                            else
                                cn = cn + "." + sName;
                            break;
                    }
                }
            }
        }

        #endregion

        #region " GetLdapPath "

        /// <summary>
        /// Get an LDAP formated path to an ou.
        /// </summary>
        /// <param name="domain"></param>
        /// <param name="ou"></param>
        /// <returns></returns>
        static public string GetLdapPath(string domain, string ou) {
            return GetLdapPath(domain, ou, "");
        }

        /// <summary>
        /// Get an LDAP formated path to a user or group cn.
        /// </summary>
        /// <param name="domain"></param>
        /// <param name="ou"></param>
        /// <param name="userOrGroup">The user or group cn.</param>
        /// <returns></returns>
        static public string GetLdapPath(string dn, string ou, string userOrGroup) {
            // Domain = domain.sd27.bc.ca
            // OU = students.sd27

            string LDAPDomain = GetLdapPath_DN(dn);
            string LDAPOU = GetLdapPath_OU(ou);
            string LDAPDomOU = "";
            string path = "";

            if(LDAPOU == "")
                LDAPDomOU = LDAPDomain;
            else
                LDAPDomOU = LDAPOU + "," + LDAPDomain;

            if(userOrGroup == "")
                path ="LDAP://" + dn + "/" + LDAPDomOU;
            else
                path ="LDAP://" + dn + "/CN=" + userOrGroup + "," + LDAPDomOU;


            return path;
        }

        /// <summary>
        /// Get the OU formatted portion of the LDAP Path
        /// </summary>
        /// <param name="ou"></param>
        /// <returns></returns>
        static public string GetLdapPath_OU(string ou) 
        {
            // OU = students.sd27
            //  string LDAPOU: "OU=students,OU=sd27"
            string[] Parts = ou.Split('.');
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

        /// <summary>
        /// Get the DC formatted portion of the LDAP Path
        /// </summary>
        /// <param name="dc"></param>
        /// <returns></returns>
        static public string GetLdapPath_DN(string dn) 
        {
            // Domain = domain.sd27.bc.ca
            //      LDAP Domain = DC=domain,DC=sd27,DC=bc,DC=ca
            string[] LDAPDomainParts = dn.Split('.');
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
