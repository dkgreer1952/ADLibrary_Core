using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.Reflection;
using System.Collections.Specialized;
using System.Collections;

namespace ADLibrary_Core;

public class GroupInfo
{
    public string GroupName { get; set; }
    public string GroupDescription { get; set; }
}
public class UserInfo
{
    public string UserID { get; set; }
    public string FullName { get; set; }
    public string Description { get; set; }
};
public class UserDetails
{
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public string MiddleInitial { get; set; }
    public string DisplayName { get; set; }
    public string Descripton { get; set; }
    public string StreetAddress { get; set; }
    public string City { get; set; }
    public string State { get; set; }
    public string ZipCode { get; set; }
    public string HomePhone { get; set; }
    public string Mobile { get; set; }
    public string EmailAddress { get; set; }
    public string WebSite { get; set; }
    public string LastLogon { get; set; }
};
public class SvcAcctInfo
{
    public string AcctName { get; set; }
    public string Description { get; set; }
    public string DisplayName { get; set; }
};

public class ADClass
{
    [DirectoryRdnPrefix("CN")]
    [DirectoryObjectClass("user")]

#pragma warning disable IDE0017 // Simplify object initialization

    private class UserPrincipalEx : UserPrincipal
    {
        UserPrincipalExSearchFilter searchFilter;
        public UserPrincipalEx(PrincipalContext context) : base(context)
        {
            searchFilter = new UserPrincipalExSearchFilter(this);
        }

        public UserPrincipalEx(PrincipalContext context, string samAccountName, string password, bool enabled)
            : base(context, samAccountName, password, enabled)
        {
            searchFilter = new UserPrincipalExSearchFilter(this);
        }


        new public UserPrincipalExSearchFilter AdvancedSearchFilter
        {
            get
            {
                if (null == searchFilter)
                    searchFilter = new UserPrincipalExSearchFilter(this);

                return searchFilter;
            }
        }

        public class UserPrincipalExSearchFilter : AdvancedFilters
        {
            public UserPrincipalExSearchFilter(Principal p) : base(p) { }
            public void LogonCount(int value, System.DirectoryServices.AccountManagement.MatchType mt)
            {
                this.AdvancedFilterSet("logonCount", value, typeof(int), mt);
            }
        }

        public static new UserPrincipalEx FindByIdentity(PrincipalContext context, string identityValue)
        {
            return (UserPrincipalEx)FindByIdentityWithType(context, typeof(UserPrincipalEx), identityValue);
        }

        public static new UserPrincipalEx FindByIdentity(PrincipalContext context, IdentityType identityType, string identityValue)
        {
            return (UserPrincipalEx)FindByIdentityWithType(context, typeof(UserPrincipalEx), identityType, identityValue);
        }

        #region custom attributes

        [DirectoryProperty("sAMAccountName")]
        public string accountName
        {
            get
            {
                if (ExtensionGet("sAMAccountName").Length != 1)
                    return string.Empty;

                return (string)ExtensionGet("sAMAccountName")[0];
            }
        }

        [DirectoryProperty("lastLogon")]
        public DateTime? lastLogon
        {

            get
            {
                if (ExtensionGet("lastLogon").Length > 0)
                {
                    var lastLogonDate = ExtensionGet("lastLogon")[0];
                    var lastLogonDateType = lastLogonDate.GetType();

                    var highPart = (Int32)lastLogonDateType.InvokeMember("HighPart", BindingFlags.GetProperty, null, lastLogonDate, null);
                    var lowPart = (Int32)lastLogonDateType.InvokeMember("LowPart", BindingFlags.GetProperty | BindingFlags.Public, null, lastLogonDate, null);

                    var longDate = ((Int64)highPart << 32 | (UInt32)lowPart);

                    DateTime dtLastLogon;

                    if (longDate > 0)
                    {
                        dtLastLogon = DateTime.FromFileTime(longDate);

                        if (base.LastLogon != null && dtLastLogon < ((DateTime)base.LastLogon).ToLocalTime())
                            dtLastLogon = ((DateTime)base.LastLogon).ToLocalTime();

                        return dtLastLogon;
                    }
                    else if (base.LastLogon != null)
                    {
                        dtLastLogon = (DateTime)base.LastLogon;
                        return dtLastLogon.ToLocalTime();
                    }
                    else
                        return null;

                }

                if (base.LastLogon == null)
                    return null;
                else
                    return ((DateTime)base.LastLogon).ToLocalTime();
            }
        }

        [DirectoryProperty("logonCount")]
        public Nullable<int> logonCount
        {
            get
            {
                if (ExtensionGet("logonCount").Length != 1)
                    return null;

                return ((Nullable<int>)ExtensionGet("logonCount")[0]);
            }
        }

        [DirectoryProperty("l")]
        public string l
        {
            get
            {
                if (ExtensionGet("l").Length != 1)
                    return string.Empty;

                return (string)ExtensionGet("l")[0];
            }
        }

        [DirectoryProperty("homePhone")]
        public string homePhone
        {
            get
            {
                if (ExtensionGet("homePhone").Length != 1)
                    return string.Empty;

                return (string)ExtensionGet("homePhone")[0];
            }
        }

        [DirectoryProperty("mobile")]
        public string mobile
        {
            get
            {
                if (ExtensionGet("mobile").Length != 1)
                    return string.Empty;

                return (string)ExtensionGet("mobile")[0];
            }
        }

        [DirectoryProperty("otherTelephone")]
        public string otherTelephone
        {
            get
            {
                if (ExtensionGet("otherTelephone").Length != 1)
                    return string.Empty;

                return (string)ExtensionGet("otherTelephone")[0];
            }
        }

        [DirectoryProperty("postalCode")]
        public string postalCode
        {
            get
            {
                if (ExtensionGet("postalCode").Length != 1)
                    return string.Empty;

                return (string)ExtensionGet("postalCode")[0];
            }
        }

        [DirectoryProperty("st")]
        public string st
        {
            get
            {
                if (ExtensionGet("st").Length != 1)
                    return string.Empty;

                return (string)ExtensionGet("st")[0];
            }
        }

        [DirectoryProperty("streetAddress")]
        public string streetAddress
        {
            get
            {
                if (ExtensionGet("streetAddress").Length != 1)
                    return string.Empty;

                return (string)ExtensionGet("streetAddress")[0];
            }
        }

        [DirectoryProperty("wWWHomePage")]
        public string wWWHomePage
        {
            get
            {
                if (ExtensionGet("wWWHomePage").Length != 1)
                    return string.Empty;

                return (string)ExtensionGet("wWWHomePage")[0];
            }
        }

        #endregion

    }

    public static GroupInfo[] GetMatchingGroups(string GroupFilter, string Domain = "")
    {
        StringCollection GroupNameColl = new();
        StringCollection GroupDescriptionColl = new ();
        SearchResultCollection resultsColl;
        GroupInfo[] Groups;

        //Default domain if not supplied
        if (Domain.Trim() == "")
            Domain = System.DirectoryServices.ActiveDirectory.Domain.GetCurrentDomain().ToString();

        if (!GroupFilter.Contains('*'))
            GroupFilter += "*";

        DirectoryContext ctx = new(DirectoryContextType.Domain, Domain);

        using (DomainController dc = DomainController.FindOne(ctx, LocatorOptions.ForceRediscovery))
        {
            using DirectorySearcher ds = dc.GetDirectorySearcher();
            //Setting PageSize causes all the domain groups to be fetched.
            //Remarkably, without this line, only 1000 records are fetched, which is equal to the
            //default value of the DirectorySearcher.SizeLimit property.

            ds.PageSize = 1000;
            //Look for security groups having scope Domain Global or Universal (Omit Local and BUILTIN groups)
            ds.Filter = string.Format("(&(objectClass=Group)(|(grouptype:1.2.840.113556.1.4.803:=-2147483646)(grouptype:1.2.840.113556.1.4.803:=-2147483640))(cn={0}))", GroupFilter);
            ds.PropertiesToLoad.Add("name");
            ds.PropertiesToLoad.Add("description");

            ds.Sort = new SortOption("name", SortDirection.Ascending);

            resultsColl = ds.FindAll();
            SearchResult[] results = new SearchResult[resultsColl.Count];
            resultsColl.CopyTo(results, 0);

            foreach (SearchResult result in results)
            {
                GroupNameColl.Add(result.Properties["name"][0].ToString());
                if (result.Properties["description"].Count > 0)
                    GroupDescriptionColl.Add(result.Properties["description"][0].ToString());
                else
                    GroupDescriptionColl.Add("NONE");
            }


            resultsColl.Dispose();
        }

        Groups = new GroupInfo[GroupNameColl.Count];

        for (int i = 0; i < GroupNameColl.Count; i++)
        {
            Groups[i] = new GroupInfo();
            Groups[i].GroupName = GroupNameColl[i].ToString();
            if (GroupDescriptionColl[i].ToString() != "NONE")
                Groups[i].GroupDescription = GroupDescriptionColl[i].ToString();
        }

        return Groups;

    }

    public static UserInfo[] GetDomainUsers(string Domain = "")
    {

        StringCollection UserIDColl = new();
        StringCollection UserNameColl = new();
        StringCollection DescriptionColl = new();
        SearchResultCollection resultsColl;
        UserInfo[] Users;

        //Default domain if not supplied
        if (Domain == null || Domain.Trim() == "")
            Domain = System.DirectoryServices.ActiveDirectory.Domain.GetCurrentDomain().ToString();

        DirectoryContext ctx = new(DirectoryContextType.Domain, Domain);

        using (DomainController dc = DomainController.FindOne(ctx, LocatorOptions.ForceRediscovery))
        {
            using DirectorySearcher ds = dc.GetDirectorySearcher();
            //Setting PageSize causes all the users to be fetched.
            //Remarkably, without this line, only 1000 records are fetched, which is equal to the
            //default value of the DirectorySearcher.SizeLimit property.

            ds.PageSize = 1000;
            ds.Filter = "(&(objectClass=user)(objectCategory=person)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"; //Find only active accounts
            ds.PropertiesToLoad.Add("samaccountname");
            ds.PropertiesToLoad.Add("givenname");
            ds.PropertiesToLoad.Add("sn");
            ds.PropertiesToLoad.Add("displayName");
            ds.PropertiesToLoad.Add("description");

            ds.Sort = new SortOption("samaccountname", SortDirection.Ascending);


            string firstName;
            string lastName;
            string fullName;

            resultsColl = ds.FindAll();
            SearchResult[] results = new SearchResult[resultsColl.Count];
            resultsColl.CopyTo(results, 0);

            foreach (SearchResult result in results)
            {
                UserIDColl.Add(result.Properties["samaccountname"][0].ToString()); //UserID

                if (result.Properties["displayName"].Count > 0)
                {
                    fullName = result.Properties["displayName"][0].ToString();
                }
                else
                {
                    if (result.Properties["givenname"].Count > 0)
                        firstName = result.Properties["givenname"][0].ToString();
                    else firstName = "";
                    if (result.Properties["sn"].Count > 0)
                        lastName = result.Properties["sn"][0].ToString();
                    else lastName = "";

                    fullName = firstName + " " + lastName;
                }

                UserNameColl.Add(fullName);

                if (result.Properties["description"].Count > 0)
                {
                    DescriptionColl.Add(result.Properties["description"][0].ToString());
                }
                else
                    DescriptionColl.Add("");
            }


            resultsColl.Dispose();
        }

        Users = new UserInfo[UserIDColl.Count];

        for (int i = 0; i < UserIDColl.Count; i++)
        {
            Users[i] = new UserInfo();
            Users[i].UserID = UserIDColl[i].ToString();
            Users[i].FullName = UserNameColl[i].ToString();
            Users[i].Description = DescriptionColl[i].ToString();
        }

        return Users;
    }

    public static GroupInfo[] GetGroupsForUser(string UserID, string Domain = "")
    {
        StringCollection GroupDescriptionColl = new();
        GroupInfo[] Groups;
        SortedList groupNameMap = new();

        //Default domain if not supplied
        if (Domain == null || Domain.Trim() == "")
            Domain = System.DirectoryServices.ActiveDirectory.Domain.GetCurrentDomain().ToString();

        PrincipalContext ctx = new(ContextType.Domain, Domain);

        UserPrincipalEx UserEx = UserPrincipalEx.FindByIdentity(ctx, IdentityType.SamAccountName, UserID);

        if ((bool)UserEx.Enabled)
        {
            DirectoryEntry domainConnection = new(); // Use this to query the default domain
                                                                    //DirectoryEntry domainConnection = new DirectoryEntry("LDAP://example.com", "username", "password"); // Use this to query a remote domain

            DirectorySearcher samSearcher = new();

            samSearcher.SearchRoot = domainConnection;
            samSearcher.Filter = "(samAccountName=" + UserEx.accountName + ")";
            samSearcher.PropertiesToLoad.Add("displayName");
            SearchResult samResult = samSearcher.FindOne();
            if (samResult != null)
            {
                DirectoryEntry theUser = samResult.GetDirectoryEntry();
                theUser.RefreshCache(new string[] { "tokenGroupsGlobalAndUniversal" });

                int groupIdx = 0;

                foreach (byte[] resultBytes in theUser.Properties["tokenGroupsGlobalAndUniversal"])
                {
                    System.Security.Principal.SecurityIdentifier mySID = new(resultBytes, 0);

                    DirectorySearcher sidSearcher = new();

                    sidSearcher.SearchRoot = domainConnection;
                    sidSearcher.Filter = "(objectSid=" + mySID.Value + ")";
                    sidSearcher.PropertiesToLoad.Add("distinguishedName");
                    sidSearcher.PropertiesToLoad.Add("name");
                    sidSearcher.PropertiesToLoad.Add("description");
                    SearchResult sidResult = sidSearcher.FindOne();
                    if (sidResult != null)
                    {
                        if (!groupNameMap.ContainsKey(sidResult.Properties["name"][0].ToString())) //Hedge against dups
                        {
                            groupNameMap[sidResult.Properties["name"][0].ToString()] = groupIdx++;

                            if (sidResult.Properties["description"].Count > 0)
                                GroupDescriptionColl.Add(sidResult.Properties["description"][0].ToString());
                            else
                                GroupDescriptionColl.Add("NONE");
                        }
                    }
                }

                Groups = new GroupInfo[groupNameMap.Count];

                int j = 0;

                foreach (DictionaryEntry groupName in groupNameMap)
                {
                    Groups[j] = new GroupInfo();
                    Groups[j].GroupName = groupName.Key.ToString();
                    if (GroupDescriptionColl[(int)groupName.Value].ToString() != "NONE")
                        Groups[j].GroupDescription = GroupDescriptionColl[(int)groupName.Value].ToString();
                    j++;
                }

                ctx.Dispose();

                return Groups;

            }
            else //No groups found for this user
            {
                Groups = new GroupInfo[1];
                Groups[0] = new GroupInfo();
                Groups[0].GroupName = "ERROR";
                Groups[0].GroupDescription = "No groups found for this user.";

                ctx.Dispose();

                return Groups;
            }

        } //end user enabled
        else
        {
            Groups = new GroupInfo[1];
            Groups[0] = new GroupInfo();
            Groups[0].GroupName = "ERROR";
            Groups[0].GroupDescription = "User account disabled";

            ctx.Dispose();

            return Groups;

        }

    }

    public static UserDetails GetUserDetails(string UserID)
    {
        if (UserID == null || UserID.Trim() == "")
            return null;

        UserDetails userDetails = new();

        PrincipalContext ctx = new(ContextType.Domain, Domain.GetCurrentDomain().ToString());

        UserPrincipalEx UserEx = UserPrincipalEx.FindByIdentity(ctx, IdentityType.SamAccountName, UserID);
        if (UserEx != null && (bool)UserEx.Enabled)
        {
            userDetails.LastName = UserEx.Surname;
            userDetails.FirstName = UserEx.GivenName;
            userDetails.DisplayName = UserEx.DisplayName;
            userDetails.HomePhone = UserEx.homePhone;
            userDetails.Mobile = UserEx.mobile;
            userDetails.EmailAddress = UserEx.EmailAddress;
            userDetails.Descripton = UserEx.Description;
            userDetails.StreetAddress = UserEx.streetAddress;
            userDetails.City = UserEx.l;
            userDetails.State = UserEx.st;
            userDetails.ZipCode = UserEx.postalCode;
            userDetails.WebSite = UserEx.wWWHomePage;
            userDetails.LastLogon = (UserEx.lastLogon != null ? UserEx.lastLogon.ToString() : "");

            UserEx.Dispose();

        } //end Active Directory data

        ctx.Dispose();

        return userDetails;
    }

    public static UserInfo[] GetDisabledUsers(string Domain = "")
    {
        StringCollection UserIDColl = new();
        StringCollection UserNameColl = new();
        SearchResultCollection resultsColl;
        UserInfo[] Users;

        //Default domain if not supplied
        if (Domain == null || Domain.Trim() == "")
            Domain = System.DirectoryServices.ActiveDirectory.Domain.GetCurrentDomain().ToString();

        DirectoryContext ctx = new(DirectoryContextType.Domain, Domain);

        using (DomainController dc = DomainController.FindOne(ctx, LocatorOptions.ForceRediscovery))
        {
            using DirectorySearcher ds = dc.GetDirectorySearcher();
            //Setting PageSize causes all the users to be fetched.
            //Remarkably, without this line, only 1000 records are fetched, which is equal to the
            //default value of the DirectorySearcher.SizeLimit property.

            ds.PageSize = 1000;
            ds.Filter = "(&(objectClass=user)(objectCategory=person)(userAccountControl:1.2.840.113556.1.4.803:=2))"; //Find disabled accounts
            ds.PropertiesToLoad.Add("name");
            ds.PropertiesToLoad.Add("samaccountname");
            ds.PropertiesToLoad.Add("givenname");
            ds.PropertiesToLoad.Add("sn");
            ds.PropertiesToLoad.Add("displayName");
            ds.PropertiesToLoad.Add("userAccountControl");

            ds.Sort = new SortOption("samaccountname", SortDirection.Ascending);


            string firstName;
            string lastName;
            string fullName;

            //int acctControl;

            resultsColl = ds.FindAll();
            SearchResult[] results = new SearchResult[resultsColl.Count];
            resultsColl.CopyTo(results, 0);

            foreach (SearchResult result in results)
            {
                UserIDColl.Add(result.Properties["samaccountname"][0].ToString()); //UserID


                if (result.Properties["displayName"].Count > 0)
                {
                    fullName = result.Properties["displayName"][0].ToString();
                }
                else
                {
                    if (result.Properties["givenname"].Count > 0)
                        firstName = result.Properties["givenname"][0].ToString();
                    else firstName = "";
                    if (result.Properties["sn"].Count > 0)
                        lastName = result.Properties["sn"][0].ToString();
                    else lastName = "";

                    fullName = firstName + " " + lastName;
                }

                UserNameColl.Add(fullName);

            }

            resultsColl.Dispose();
        }

        Users = new UserInfo[UserIDColl.Count];

        for (int i = 0; i < UserIDColl.Count; i++)
        {
            Users[i] = new UserInfo();
            Users[i].UserID = UserIDColl[i].ToString();
            Users[i].FullName = UserNameColl[i].ToString();
        }


        return Users;


    }

    public static UserInfo[] GetLockedUsers(string Domain = "")
    {
        StringCollection UserIDColl = new();
        StringCollection UserNameColl = new();
        SearchResultCollection resultsColl;
        UserInfo[] Users;

        //Default domain if not supplied
        if (Domain == null || Domain.Trim() == "")
            Domain = System.DirectoryServices.ActiveDirectory.Domain.GetCurrentDomain().ToString();

        DirectoryContext ctx = new(DirectoryContextType.Domain, Domain);

        PrincipalContext pCtx = new(ContextType.Domain, Domain);


        using (DomainController dc = DomainController.FindOne(ctx, LocatorOptions.ForceRediscovery))
        {
            using DirectorySearcher ds = dc.GetDirectorySearcher();
            //Setting PageSize causes all the users to be fetched.
            //Remarkably, without this line, only 1000 records are fetched, which is equal to the
            //default value of the DirectorySearcher.SizeLimit property.

            //First, identify a list of potentially locked accounts.  Some on this list may not
            //actually be locked, but rather have their password reset, and awaiting login to
            //clear the lockoutTime flag.

            ds.PageSize = 1000;
            ds.Filter = "(&(objectClass=user)(objectCategory=person)(lockoutTime>=1))";
            ds.PropertiesToLoad.Add("name");
            ds.PropertiesToLoad.Add("samaccountname");
            ds.PropertiesToLoad.Add("givenname");
            ds.PropertiesToLoad.Add("sn");
            ds.PropertiesToLoad.Add("displayName");
            ds.PropertiesToLoad.Add("userAccountControl");

            ds.Sort = new SortOption("samaccountname", SortDirection.Ascending);


            string firstName;
            string lastName;
            string fullName;
            string userID;

            //int acctControl;

            resultsColl = ds.FindAll();
            SearchResult[] results = new SearchResult[resultsColl.Count];
            resultsColl.CopyTo(results, 0);

            //Now cull down the list of candidates by using the UserPrincipal class, which
            //attempts to calculate whether the account is actually locked.
            //Note that the classes used here are only supported in .NET 3.5

            foreach (SearchResult result in results)
            {
                userID = result.Properties["samaccountname"][0].ToString();
                UserPrincipal userPrincipal = new(pCtx);

                userPrincipal.SamAccountName = userID;

                PrincipalSearcher ps = new(userPrincipal);

                userPrincipal = (UserPrincipal)ps.FindOne();

                ps.Dispose();

                bool bIsLocked = false;

                if (userPrincipal != null)
                    bIsLocked = userPrincipal.IsAccountLockedOut();
                userPrincipal.Dispose();
                if (bIsLocked)
                {

                    UserIDColl.Add(userID); //UserID


                    if (result.Properties["displayName"].Count > 0)
                    {
                        fullName = result.Properties["displayName"][0].ToString();
                    }
                    else
                    {
                        if (result.Properties["givenname"].Count > 0)
                            firstName = result.Properties["givenname"][0].ToString();
                        else firstName = "";
                        if (result.Properties["sn"].Count > 0)
                            lastName = result.Properties["sn"][0].ToString();
                        else lastName = "";

                        fullName = firstName + " " + lastName;
                    }

                    UserNameColl.Add(fullName);
                }

            }


            resultsColl.Dispose();
        }

        Users = new UserInfo[UserIDColl.Count];

        for (int i = 0; i < UserIDColl.Count; i++)
        {
            Users[i] = new UserInfo();
            Users[i].UserID = UserIDColl[i].ToString();
            Users[i].FullName = UserNameColl[i].ToString();
        }


        return Users;


    }

    public static UserInfo[] GetMatchingUsers(string LastName = "", string FirstName = "", string Domain = "")
    {
        StringCollection UserIDColl = new();
        StringCollection UserNameColl = new();
        SearchResultCollection resultsColl;
        UserInfo[] Users;

        if (LastName == null) LastName = "";
        if (FirstName == null) FirstName = "";

        if (!FirstName.Contains('*'))
            FirstName += "*";

        if (!LastName.Contains('*'))
            LastName += "*";

        //Default domain if not supplied
        if ((Domain == null) || Domain.Trim() == "")
            Domain = System.DirectoryServices.ActiveDirectory.Domain.GetCurrentDomain().ToString();

        DirectoryContext ctx = new(DirectoryContextType.Domain, Domain);

        using (DomainController dc = DomainController.FindOne(ctx, LocatorOptions.ForceRediscovery))
        {
            using DirectorySearcher ds = dc.GetDirectorySearcher();
            //Setting PageSize causes all the users to be fetched.
            //Remarkably, without this line, only 1000 records are fetched, which is equal to the
            //default value of the DirectorySearcher.SizeLimit property.

            ds.PageSize = 1000;
            ds.Filter = string.Format("(&(objectClass=user)(objectCategory=person)(sn={0})(givenname={1})(!(userAccountControl:1.2.840.113556.1.4.803:=2)))", LastName, FirstName); //Find only active accounts
            ds.PropertiesToLoad.Add("samaccountname");
            ds.PropertiesToLoad.Add("givenname");
            ds.PropertiesToLoad.Add("sn");
            ds.PropertiesToLoad.Add("displayName");

            ds.Sort = new SortOption("samaccountname", SortDirection.Ascending);


            string firstName;
            string lastName;
            string fullName;

            resultsColl = ds.FindAll();
            SearchResult[] results = new SearchResult[resultsColl.Count];
            resultsColl.CopyTo(results, 0);

            foreach (SearchResult result in results)
            {

                UserIDColl.Add(result.Properties["samaccountname"][0].ToString()); //UserID


                if (result.Properties["displayName"].Count > 0)
                {
                    fullName = result.Properties["displayName"][0].ToString();
                }
                else
                {
                    if (result.Properties["givenname"].Count > 0)
                        firstName = result.Properties["givenname"][0].ToString();
                    else firstName = "";
                    if (result.Properties["sn"].Count > 0)
                        lastName = result.Properties["sn"][0].ToString();
                    else lastName = "";

                    fullName = firstName + " " + lastName;
                }

                UserNameColl.Add(fullName);

            }


            resultsColl.Dispose();
        }

        Users = new UserInfo[UserIDColl.Count];

        for (int i = 0; i < UserIDColl.Count; i++)
        {
            Users[i] = new UserInfo();
            Users[i].UserID = UserIDColl[i].ToString();
            Users[i].FullName = UserNameColl[i].ToString();
        }


        return Users;

    }

    public static UserInfo[] GetGroupMembers(string GroupName, string Domain = "")
    {

        SortedList userIDMap = new();
        StringCollection UserNameColl = new();
        StringCollection DescriptionColl = new();
        UserInfo[] Users;

        if (GroupName == null || GroupName.Trim() == "")
            return null;

        //Default domain if not supplied
        if (Domain == null || Domain.Trim() == "")
            Domain = System.DirectoryServices.ActiveDirectory.Domain.GetCurrentDomain().ToString();


        PrincipalContext ctx = new(ContextType.Domain, Domain);
        GroupPrincipal group;

        try
        {
            group = GroupPrincipal.FindByIdentity(ctx, GroupName);

            if (group == null)
            {
                Users = new UserInfo[1];
                Users[0] = new UserInfo();
                Users[0].UserID = "ERROR";
                Users[0].FullName = "Group name not found.";
                return Users;
            }

            // get the users for the group principal and store the results in a PrincipalSearchResult object
            PrincipalSearchResult<Principal> results = group.GetMembers(true); //Recursive nested group search

            var itResults = results.GetEnumerator();

            int j = 0;

            using (itResults)
            {
                while (itResults.MoveNext())
                {
                    try
                    {
                        if (itResults.Current.GetType() == typeof(UserPrincipal))  //Skip ComputerPrincipals
                        {
                            UserPrincipal p = (UserPrincipal)itResults.Current;

                            UserPrincipalEx UserEx = UserPrincipalEx.FindByIdentity(ctx, IdentityType.SamAccountName, p.SamAccountName);
                            if (!userIDMap.ContainsKey(p.SamAccountName) && (bool)UserEx.Enabled)
                            {
                                userIDMap[p.SamAccountName] = j++;
                                if (p.DisplayName != null && p.DisplayName != "")
                                    UserNameColl.Add(p.DisplayName);
                                else
                                    UserNameColl.Add("");  //To test for below

                                if (p.Description != null && p.Description != "")
                                    DescriptionColl.Add(p.Description);
                                else
                                    DescriptionColl.Add("");

                            }
                        }

                    }
                    catch (NoMatchingPrincipalException)
                    {
                        //Swallow these. 
                        //They are extraneously thrown when orphaned SIDs are encountered.
                        //http://stackoverflow.com/questions/16221974/getauthorizationgroups-is-throwing-exception
                        //http://social.msdn.microsoft.com/Forums/vstudio/en-US/9dd81553-3539-4281-addd-3eb75e6e4d5d/getauthorizationgroups-fails-with-nomatchingprincipalexception

                        continue; //while
                    }

                }
            }
        }
        catch (MultipleMatchesException)
        {
        }


        Users = new UserInfo[userIDMap.Count];

        int u = 0;

        foreach (DictionaryEntry userID in userIDMap)
        {
            Users[u] = new UserInfo();
            Users[u].UserID = userID.Key.ToString();
            Users[u].FullName = UserNameColl[(int)userID.Value].ToString();
            Users[u].Description = DescriptionColl[(int)userID.Value].ToString();
            u++;
        }

        return Users;

    } //GetGroupMembers

    public static GroupInfo[] GetGroupSubGroups(string GroupName, string Domain = "")
    {

        GroupInfo[] Groups;
        SortedList groupNameMap = new();
        StringCollection GroupDescriptionColl = new();

        if (GroupName == null || GroupName.Trim() == "")
            return null;

        //Default domain if not supplied
        if (Domain == null || Domain.Trim() == "")
            Domain = System.DirectoryServices.ActiveDirectory.Domain.GetCurrentDomain().ToString();

        PrincipalContext ctx = new(ContextType.Domain, Domain);

        try
        {
            GroupPrincipal group = GroupPrincipal.FindByIdentity(ctx, GroupName);

            if (group == null)
            {
                Groups = new GroupInfo[1];
                Groups[0] = new GroupInfo();
                Groups[0].GroupName = "ERROR";
                Groups[0].GroupDescription = "Group name not found.";
                return Groups;
            }

            // get the users for the group principal and store the results in a PrincipalSearchResult object
            PrincipalSearchResult<Principal> results = group.GetMembers(false); //Recursive nested group search

            var itResults = results.GetEnumerator();

            int j = 0;

            using (itResults)
            {
                while (itResults.MoveNext())
                {
                    try
                    {
                        Principal p = (Principal)itResults.Current;

                        if (p.StructuralObjectClass.ToLower() == "group") //Return only subgroups
                        {
                            GroupPrincipal gp = (GroupPrincipal)p;
                            if (gp.GroupScope != GroupScope.Local)
                            {

                                if (!groupNameMap.ContainsKey(gp.Name)) //Hedge against dupes
                                {
                                    groupNameMap[gp.Name] = j++;
                                    if (gp.Description != null && gp.Description != "")
                                        GroupDescriptionColl.Add(gp.Description);
                                    else
                                        GroupDescriptionColl.Add("NONE");  //To test for below
                                }
                            }
                        }

                    }
                    catch (NoMatchingPrincipalException)
                    {
                        //Swallow these. 
                        //They are extraneously thrown when orphaned SIDs are encountered.
                        //http://stackoverflow.com/questions/16221974/getauthorizationgroups-is-throwing-exception
                        //http://social.msdn.microsoft.com/Forums/vstudio/en-US/9dd81553-3539-4281-addd-3eb75e6e4d5d/getauthorizationgroups-fails-with-nomatchingprincipalexception

                        continue; //while
                    }

                }
            }

            Groups = new GroupInfo[groupNameMap.Count];

            int g = 0;

            foreach (DictionaryEntry groupName in groupNameMap)
            {
                Groups[g] = new GroupInfo();
                Groups[g].GroupName = groupName.Key.ToString();
                if (GroupDescriptionColl[(int)groupName.Value].ToString() != "NONE")
                    Groups[g].GroupDescription = GroupDescriptionColl[(int)groupName.Value].ToString();
                g++;
            }
            return Groups;
        }
        catch (MultipleMatchesException)
        {
            Groups = Array.Empty<GroupInfo>();
            return Groups;
        }


    } //GetGroupSubGroups

    public static string[] GetDomainControllers(string DomainName = "")
    {
        string[] DC;

        StringCollection DC_Coll = new();

        //Default domain if not supplied
        if (DomainName == null || DomainName.Trim() == "")
            DomainName = Domain.GetCurrentDomain().ToString();

        DirectoryContext ctx = new(DirectoryContextType.Domain, DomainName);

        Domain domain = Domain.GetDomain(ctx);

        foreach (DomainController dc in domain.DomainControllers)
            DC_Coll.Add(dc.ToString());

        DC = new string[DC_Coll.Count];

        for (int i = 0; i < DC_Coll.Count; i++)
            DC[i] = DC_Coll[i].ToString();
        return DC;
    }

    public static bool IsUserInGroup(string UserID, string Group, string Domain = "")
    {
        bool bUserInGroup = false;

        if (UserID == null || UserID.Trim() == "" || Group == null || Group.Trim() == "")
            return false;

        //Default domain if not supplied
        if (Domain == null || Domain.Trim() == "")
            Domain = System.DirectoryServices.ActiveDirectory.Domain.GetCurrentDomain().ToString();

        PrincipalContext pCtx = new(ContextType.Domain, Domain);

        UserPrincipalEx UserEx = UserPrincipalEx.FindByIdentity(pCtx, IdentityType.SamAccountName, UserID);
        if (UserEx == null || !(bool)UserEx.Enabled) return false;
        GroupPrincipal gp = new(pCtx);
        gp.Name = Group;

        PrincipalSearcher ps = new(gp);
        gp = (GroupPrincipal)ps.FindOne();
        ps.Dispose();

        // An attempt was made here to use: UserEx.IsMemberOf(gp), however it erroneously returns false in the case
        //where the group coincides with the users primary group, which is normally Domain Users, but could be some other group.
        //Instead, we resort here to iterating over the users groups, which includes their primary group, looking for a match.

        if (gp != null)
        {
            PrincipalSearchResult<Principal> gpresults = UserEx.GetGroups();  //This returns all groups, including the primary.

            var itResults = gpresults.GetEnumerator();

            while (itResults.MoveNext())
            {
                if (itResults.Current.GetType() == typeof(GroupPrincipal))
                {
                    if (itResults.Current.Equals(gp))
                    {
                        bUserInGroup = true;
                        break;
                    }
                }
            }

            gp.Dispose();
        }

        return bUserInGroup;
    }

}
