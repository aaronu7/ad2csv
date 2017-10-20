using NUnit.Framework;
using System.Reflection;
using System.IO;
using System.Data;
using System.Collections;
using System;

using ad2csv.Db.LDAP;

namespace ad2csv_test.UnitTests
{
    [TestFixture]
    public class BasicTests
    {
        [SetUp] public void Setup() {}
        [TearDown] public void TestTearDown() {}

        #region " TestADQuery " 

        public static IEnumerable InputTestADQuery {
            get {
                string domain = "domain.bc.ca";
                string usr = "administrator";
                string pwd = "";
                string ouLimit = "";
                int pageSize = 2000;
                int pageSizeTestLimit = 100000; // set this to the absolute maximum you would ever expect.. and then a bit more.

                // Test users
                //yield return new TestCaseData(domain, usr, pwd, "user", ouLimit, pageSize, pageSizeTestLimit); 
                //yield return new TestCaseData(domain, usr, pwd, "group", ouLimit, pageSize, pageSizeTestLimit);
                //yield return new TestCaseData(domain, usr, pwd, "organizationalunit", ouLimit, pageSize, pageSizeTestLimit); 
                yield return new TestCaseData(domain, usr, pwd, "", ouLimit, pageSize, pageSizeTestLimit); 
            }
        }

        //[Ignore("")]
        [Test]
        [TestCaseSource("InputTestADQuery")]
        public void TestADQuery(string domain, string usr, string pwd, string objectClass, string ouLimit, int pageSize, int pageSizeTestLimit) {

            // Get the DataTable extracted
            DataTable oDt = null;
            if (objectClass.ToLower() == "user") {
                oDt = Ad2DataTable.ExtractDataTable_Users(domain, usr, pwd, ouLimit, pageSize);
            } else if (objectClass.ToLower() == "organizationalunit") {
                oDt = Ad2DataTable.ExtractDataTable_Units(domain, usr, pwd, ouLimit, pageSize);
            } else if (objectClass.ToLower() == "group") {
                oDt = Ad2DataTable.ExtractDataTable_Groups(domain, usr, pwd, ouLimit, pageSize);
            }
            Assert.NotNull(oDt, objectClass + " DataTable is null. These tests may be disabled or there is a possible connectivity issue, check your test settings in BasicTests.cs.");

            // Get a simple query count
            Int32 queryCount = HelperLdapGet.GetObjectSetCount(domain, ouLimit, usr, pwd, objectClass, "", pageSizeTestLimit);
            Assert.AreEqual(oDt.Rows.Count, queryCount, objectClass + " DataTable count does not match the actual query count. You may need to increase your pageSize.");
        }

        #endregion

        #region " TestGetLdapPathParts " 

        public static IEnumerable InputTestGetLdapPathParts {
            get {

                yield return new TestCaseData(
                    "LDAP://root.sd27.bc.ca/CN=001,OU=numerics,OU=students,OU=sd27,DC=root,DC=sd27,DC=bc,DC=ca", 
                    "LDAP://root.sd27.bc.ca/",
                    "root.sd27.bc.ca", "numerics.students.sd27", "001"); 
            }
        }

        //[Ignore("")]
        [Test]
        [TestCaseSource("InputTestGetLdapPathParts")]
        public void TestGetLdapPathParts(string ldapPath, string ldapRoot, string expectedDN, string expectedOU, string expectedCN) {
            string dn = "";
            string ou = "";
            string cn = "";
            HelperLdapGet.GetLdapPathParts(ldapPath, ldapRoot, ref cn, ref ou, ref dn);

            Assert.AreEqual(expectedDN, dn, "DN is not as expected. Expected: " + expectedDN + "  Actual: " + dn);
            Assert.AreEqual(expectedOU, ou, "OU is not as expected. Expected: " + expectedOU + "  Actual: " + ou);
            Assert.AreEqual(expectedCN, cn, "CN is not as expected. Expected: " + expectedCN + "  Actual: " + cn);
        }

        #endregion

        #region " TestGetLdapPath " 

        public static IEnumerable InputTestGetLdapPath {
            get {

                yield return new TestCaseData(
                    "root.sd27.bc.ca", "numerics.students.sd27", "001",
                    "LDAP://root.sd27.bc.ca/CN=001,OU=numerics,OU=students,OU=sd27,DC=root,DC=sd27,DC=bc,DC=ca"); 
            }
        }

        //[Ignore("")]
        [Test]
        [TestCaseSource("InputTestGetLdapPath")]
        public void TestGetLdapPath(string domain, string ou, string cn, string expectedLDapPath) {
            string ldapPath = HelperLdapGet.GetLdapPath(domain, ou, cn);
            Assert.AreEqual(expectedLDapPath, ldapPath, "Path is not as expected. Expected: " + expectedLDapPath + "  Actual: " + ldapPath);
        }

        #endregion

        #region " TestGetObjectNameQuery " 

        public static IEnumerable InputTestGetObjectNameQuery {
            get {

                yield return new TestCaseData("user", "john.doe", "cn=john.doe");
                yield return new TestCaseData("group", "staff_readwrite", "cn=staff_readwrite");
                yield return new TestCaseData("organizationalunit", "staff", "ou=staff");
            }
        }

        //[Ignore("")]
        [Test]
        [TestCaseSource("InputTestGetObjectNameQuery")]
        public void TestGetObjectNameQuery(string objectClass, string cn, string expectedQuery) {
            string query = HelperLdapGet.GetObjectNameQuery(objectClass, cn);
            Assert.AreEqual(expectedQuery, query, "Query is not as expected. Expected: " + expectedQuery + "  Actual: " + query);
        }

        #endregion
    }
}
  