using Avanade.LAM.CollabBOT.LAM;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.Online.SharePoint.TenantManagement;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Sharing;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using SharePointPnP;


namespace CollabLAMBot.LAM
{
    public class SharePointPrimary
    {
        string SPOAdmin = "labani@AvaIndCollabSL.onmicrosoft.com";
        string SPOAdminPassowrd = "$D9920526814l";
        string SPOAdminURL = "https://avaindcollabsl-admin.sharepoint.com/";
        string clientID = "b6c760b4-defd-4471-9ab0-15adbfd4b1a7";
        string clientSecret = "agRvUdPm0ujJxPc7qRshc24xJY64DHzwEC8jv7tADyI=";
        string siteCollectionURL = string.Empty;

        /// <summary>
        /// constructor
        /// </summary>
        /// <param name="siteCollectionURL"></param>
        public SharePointPrimary(string siteCollectionURL)
        {
            this.siteCollectionURL = siteCollectionURL;
        }
        /// <summary>
        /// constrcutor
        /// </summary>
        public SharePointPrimary() { }

        public bool IsSiteCollectionStorageQuotaUpdated(string _strURL, int _intNewQuota)
        {
            bool isStorageUpdated = false;
            string siteCollectionUrl = SPOAdminURL;

            try
            {
                ClientContext clientContext;
                using (clientContext = GetClientContext(siteCollectionUrl, clientID, clientSecret))
                {
                    clientContext.ExecuteQuery();
                    Tenant currentO365Tenant = new Tenant(clientContext);
                    clientContext.ExecuteQuery();

                    SiteProperties propertyColl = currentO365Tenant.GetSitePropertiesByUrl(_strURL,true);

                    if (propertyColl != null)
                    {
                        clientContext.Load(propertyColl);
                        clientContext.ExecuteQuery();

                        //propertyColl.Title += "_storage Updated By Bot";
                        propertyColl.StorageMaximumLevel += (_intNewQuota * 1024);                        
                        propertyColl.Update();

                        clientContext.Load(propertyColl);
                        clientContext.ExecuteQuery();
                        isStorageUpdated = true;
                    }
                }
            }
            catch(Exception ex)
            {
                isStorageUpdated = false;
            }

            return isStorageUpdated;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="url"></param>
        /// <param name="clientid"></param>
        /// <param name="clientsecret"></param>
        /// <returns></returns>
        public static ClientContext GetClientContext(string url, string clientid, string clientsecret)
        {
            Uri siteUri = new Uri(url);
            //Get the realm for the URL  
            string realm = TokenHelper.GetRealmFromTargetUrl(siteUri);
            //Get the access token for the URL.   
            //  Requires this app to be registered with the tenant  
            string accessToken = TokenHelper.GetAppOnlyAccessTokenCustom(
              TokenHelper.SharePointPrincipal,
              siteUri.Authority, realm, clientid, clientsecret).AccessToken;
            var clientContext = TokenHelper.GetClientContextWithAccessToken(siteUri.ToString(), accessToken);
            return clientContext;
        }

        public async Task<bool> IsSiteCollectionCreated(string _strSiteTitle, string _strPrimaryAdmin)
        {
            bool isSiteCreated = false;
            string status = string.Empty;
            string siteCollectionUrl = SPOAdminURL;
            try
            {
                ClientContext clientContext;
                using (clientContext = GetClientContext(siteCollectionUrl, clientID, clientSecret))
                {
                    clientContext.ExecuteQuery();
                    Tenant currentO365Tenant = new Tenant(clientContext);
                    clientContext.ExecuteQuery();

                    /*make UserCodeMaximumLevel 1 once license is restored*/
                    var newsite = new SiteCreationProperties()
                    {
                        Url = Constants.RootSiteCollectionURL+Constants.ManagedPath+"SOHA_"+_strSiteTitle,
                        Owner = _strPrimaryAdmin,
                        Template = "STS#0",
                        StorageMaximumLevel = 2,
                        UserCodeMaximumLevel = 0, 
                        UserCodeWarningLevel = 0, 
                        StorageWarningLevel = 1,
                        Title = "SOHA_"+_strSiteTitle,
                        CompatibilityLevel = 15,                        

                    };

                    SpoOperation spo = currentO365Tenant.CreateSite(newsite);
                    clientContext.Load(currentO365Tenant);
                    //Load IsComplete property to check if the provisioning of the Site Collection is complete.
                    clientContext.Load(spo, i => i.IsComplete);
                    try
                    {
                        clientContext.ExecuteQuery();

                        while (!spo.IsComplete)
                        {
                            //Wait and then try again
                            //System.Threading.Thread.Sleep(2000);
                            await Task.Delay(10000);
                            spo.RefreshLoad();
                            try
                            {
                                isSiteCreated = true;
                            }
                            catch (Exception ex)
                            {
                                isSiteCreated = false;
                            }
                        }

                        isSiteCreated = true;
                    }
                    catch(Exception ex)
                    {
                        isSiteCreated = false;
                        status = "Error in site creation " + ex.Message;
                    }
                }
            }
            catch(Exception ex)
            {
                isSiteCreated = false;
                status = "Error in site creation " + ex.Message;
            }

            return isSiteCreated;
        }

        public bool DoesContainSpecialCharacter(string _inputSiteCollectionTitle)
        {
            bool isValid = false;
            var regexItem = new Regex("^[a-zA-Z0-9 _]*$");

            if (regexItem.IsMatch(_inputSiteCollectionTitle))
            {
                isValid = true;
            }

            return isValid;

        }

        public bool DoesURLExist(string _inputSiteCollectionTitle)
        {
            bool isSiteCollectionURLExisting = false;
            string siteCollectionUrl = SPOAdminURL;


            string _inputSiteCollectionURL = Constants.RootSiteCollectionURL + Constants.ManagedPath + _inputSiteCollectionTitle;

            try
            {
                ClientContext clientContext;
                using (clientContext = GetClientContext(siteCollectionUrl, clientID, clientSecret))
                {
                    clientContext.ExecuteQuery();
                    Tenant currentO365Tenant = new Tenant(clientContext);
                    clientContext.ExecuteQuery();

                    SPOSitePropertiesEnumerable sitePropEnumerable = currentO365Tenant.GetSiteProperties(0, true);
                    clientContext.Load(sitePropEnumerable);
                    clientContext.ExecuteQuery();

                    List<SiteProperties> sitePropertyCollection = new List<SiteProperties>();
                    sitePropertyCollection.AddRange(sitePropEnumerable);

                    SiteProperties property1 = sitePropertyCollection.Find(s => s.Url.ToLower().Equals(_inputSiteCollectionURL.ToLower())) ;
                    isSiteCollectionURLExisting = property1 != null ? true : false;                    

                }
            }
            catch(Exception ex)
            {
                isSiteCollectionURLExisting = false;
            }
            return isSiteCollectionURLExisting;
        }

        public /*Dictionary<string, string>*/ List<ListItem> SearchHelpArticleonTopic(string _strhelpArticle)
        {
            string siteCollectionUrl = "https://avaindcollabsl.sharepoint.com/sites/SOHA_HelpRepository/";
            Dictionary<string, string> searchedDocs = new Dictionary<string, string>();

           
            List<ListItem> currentSearchTopics = new List<ListItem>();

            try
            {
                ClientContext clientContext;
                using (clientContext = GetClientContext(siteCollectionUrl, clientID, clientSecret))
                {
                    clientContext.ExecuteQuery();
                    Web web = clientContext.Web;
                    clientContext.Load(web);
                    List articleRepository = web.Lists.GetByTitle("HelpRepository");
                    clientContext.Load(articleRepository);
                    clientContext.ExecuteQuery();
                    CamlQuery query = new CamlQuery();
                    CamlQuery camlQuery = new CamlQuery();                  

                    camlQuery.ViewXml = "<View Scope=\"RecursiveAll\">"
                                + "<ViewFields><FieldRef Name=\"Title\"/>"
                                + "<FieldRef Name=\"Tags\"/>"
                                + "<FieldRef Name=\"Description\"/>"
                                + "<FieldRef Name=\"ThumbnailURL\"/>"
                                + "<FieldRef Name=\"Modified\"/></ViewFields>"
                                + "<RowLimit>500</RowLimit></View>";

                    ListItemCollection articleCollection = articleRepository.GetItems(query);                   
                    clientContext.Load(articleCollection,
                        articles => articles.Include(
                            a => a.Id,
                            a => a.FieldValuesAsHtml,
                            a => a.DisplayName,
                            a => a.File));
                    clientContext.ExecuteQuery();

                    string test = articleCollection[0].File.ServerRelativeUrl;

                    List<ListItem> lstListItemCollection = new List<ListItem>();
                    lstListItemCollection.AddRange(articleCollection);

                    
                    currentSearchTopics = lstListItemCollection.FindAll(a => Convert.ToString(a.FieldValuesAsHtml["Tags"]).Contains(_strhelpArticle.ToLower()));
                    if (currentSearchTopics != null && currentSearchTopics.Count > 0)
                    {
                        foreach (ListItem eachFoundArticle in currentSearchTopics)
                        {
                            searchedDocs.Add(eachFoundArticle.DisplayName, Constants.RootSiteCollectionURL+eachFoundArticle.File.ServerRelativeUrl);                          
                        }
                    }
                }
            }
            catch(Exception ex)
            {

            }
            return currentSearchTopics;

        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool IsValidSiteCollectionURL()
        {
            bool isValid = true;
            string message = string.Empty;
            string siteCollectionUrl = this.siteCollectionURL;

            try
            {
                Uri uriResult;
                bool isUrlValid = Uri.TryCreate(siteCollectionUrl, UriKind.Absolute, out uriResult) && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
                if (!isUrlValid)
                {
                    message =  "This is not a valid Uri.";
                    isValid = false;
                }
                else
                {
                    ClientContext clientContext;
                    using (clientContext = GetClientContext(siteCollectionUrl, clientID, clientSecret))
                    {
                        try
                        {
                            clientContext.ExecuteQuery();
                        }
                        catch (Exception ex)
                        {
                            isValid = false;
                            message =  "Please recheck the url. It's not valid in our tenant . " + ex.Message;

                        }
                    }
                }
            }
            catch(Exception ex)
            {
                isValid = false;
                message = "Unexpected error occurred." + ex.Message;
            }
            return isValid;
        }

        public bool HasPermissionGrantedToUser(string _strURL, string _strUserID, int _intRoleType)
        {
            string message = string.Empty;
            bool isPermissionGranted = false;
            try
            {
                ClientContext clientContext;
                using (clientContext = GetClientContext(_strURL, clientID, clientSecret))
                {
                    try
                    {
                        clientContext.ExecuteQuery();

                        Web web = clientContext.Web;
                        clientContext.Load(web);
                        clientContext.Load(web.AllProperties);                        
                        clientContext.ExecuteQuery();
                        var roleDefinitions = web.RoleDefinitions;
                        clientContext.ExecuteQuery();                        

                        User oUser = web.EnsureUser(_strUserID);
                        clientContext.Load(oUser);
                        clientContext.ExecuteQuery();

                        RoleDefinition _inputRoleDefinition = null;

                        switch (_intRoleType)
                        {
                            case 0:
                                 _inputRoleDefinition = roleDefinitions.GetByType(RoleType.None);
                                break;

                            case 1:
                                _inputRoleDefinition = roleDefinitions.GetByType(RoleType.Guest);
                                break;

                            case 2:
                                _inputRoleDefinition = roleDefinitions.GetByType(RoleType.Reader);
                                break;

                            case 3:
                                _inputRoleDefinition = roleDefinitions.GetByType(RoleType.Contributor);
                                break;

                            case 4:
                                _inputRoleDefinition = roleDefinitions.GetByType(RoleType.WebDesigner);
                                break;

                            case 5:
                                _inputRoleDefinition = roleDefinitions.GetByType(RoleType.Administrator);
                                break;

                            case 6:
                                _inputRoleDefinition = roleDefinitions.GetByType(RoleType.Editor);
                                break;

                            case 7:
                                _inputRoleDefinition = roleDefinitions.GetByType(RoleType.Reviewer);
                                break;

                            default:
                                _inputRoleDefinition = roleDefinitions.GetByType(RoleType.Reader);
                                break;
                        }

                        if (_inputRoleDefinition != null)
                        {

                            clientContext.Load(_inputRoleDefinition);
                            clientContext.ExecuteQuery();

                            web.RoleAssignments.Add(web.SiteUsers.GetById(oUser.Id),
                                new RoleDefinitionBindingCollection(clientContext) { _inputRoleDefinition });
                            clientContext.ExecuteQuery();
                            isPermissionGranted = true;
                        }                       
                    }
                    catch (Exception ex)
                    {
                        isPermissionGranted = false;
                        message = "Error in assigning permission . " + ex.Message;

                    }
                }
            }
            catch(Exception ex)
            {
                isPermissionGranted = false;
            }

            return isPermissionGranted;

        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool IsValidSPOUser(string _inputSPOUserID)
        {
            bool isValid = false;
            string message = string.Empty;
            string siteCollectionUrl = SPOAdminURL;
            List<User> userCollection = new List<User>();

            try
            {
                ClientContext clientContext;
                using (clientContext = GetClientContext(siteCollectionUrl, clientID, clientSecret))
                {
                    try
                    {
                        clientContext.ExecuteQuery();
                        Web oWeb = clientContext.Site.RootWeb;
                        clientContext.Load(oWeb);
                        try
                        {
                            clientContext.ExecuteQuery();

                            PeopleManager peopleManager = new PeopleManager(clientContext);
                            UserCollection users = clientContext.Web.SiteUsers;
                            clientContext.Load(users);
                            try
                            {
                                clientContext.ExecuteQuery();
                                userCollection.AddRange(users);

                                User matchedUser = userCollection.Find(u => u.Email.ToLower().Equals(_inputSPOUserID.ToLower()));
                                if (matchedUser != null)
                                {
                                    PersonProperties personProperties = peopleManager.GetPropertiesFor(matchedUser.LoginName);
                                    clientContext.Load(personProperties, p => p.UserProfileProperties, p => p.PersonalUrl);
                                    try
                                    {
                                        clientContext.ExecuteQuery();
                                        isValid = true;
                                    }
                                    catch (Exception ex)
                                    {
                                        isValid = false;
                                        message = "Error in loading user profile properties for this user." + ex.Message;
                                    }
                                }                              
                            }
                                catch (Exception ex)
                            {
                                isValid = false;
                                message = "Error in loading users." + ex.Message;
                            }
                        }
                        catch(Exception ex)
                        {
                            isValid = false;
                            message = "Error in loading current web. " + ex.Message;
                        }


                    }
                    catch (Exception ex)
                    {
                        isValid = false;
                        message = "Please recheck the url. It's not valid in our tenant . " + ex.Message;

                    }
                }

            }
            catch(Exception ex)
            {
                isValid = false;
                message = "Unexpected error occurred." + ex.Message;
            }

            return isValid;
        }

        public bool IsSiteCollectionAdmin(string _inputSPOUserID)
        {
            bool isSiteAdmin = false;
            List<string> lstSiteCollectionAdmins = new List<string>();

            lstSiteCollectionAdmins = GetSiteCollectionAdmins();
            if (lstSiteCollectionAdmins.Count > 0)
            {
                if (lstSiteCollectionAdmins.Contains(_inputSPOUserID.ToLower()))
                    isSiteAdmin = true;
                else
                    isSiteAdmin = false;
            }
            return isSiteAdmin;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="siteCollectionUrl"></param>
        /// <param name="_inputSPOUserID"></param>
        public string CheckUserPermission(string siteCollectionUrl, string _inputSPOUserID)
        {
            string message = string.Empty;
            string _userPermissions = string.Empty;
            try
            {
                ClientContext clientContext =null ;
                Web oWeb = null;
                using (clientContext = GetClientContext(siteCollectionUrl, clientID, clientSecret))
                {
                    try
                    {
                        clientContext.ExecuteQuery();
                        oWeb = clientContext.Site.RootWeb;
                        clientContext.Load(oWeb);

                        clientContext.ExecuteQuery();
                        User oUser = oWeb.EnsureUser(_inputSPOUserID);
                        clientContext.Load(oUser);

                        try
                        {
                            clientContext.ExecuteQuery();

                            var permissions = oWeb.GetUserEffectivePermissions(oUser.LoginName);                            
                            try
                            {
                                clientContext.ExecuteQuery();
                                bool test = permissions.Value.Has(PermissionKind.ManagePermissions);

                                foreach(var permissionKind in Enum.GetValues(typeof(PermissionKind)).Cast<PermissionKind>())
                                {
                                    if (permissionKind != PermissionKind.EmptyMask)
                                    {
                                        var permission = Enum.GetName(typeof(PermissionKind), permissionKind);
                                        var hasPermission = permissions.Value.Has(permissionKind);
                                        if (hasPermission)
                                        {
                                            var words =
                                                    Regex.Matches(permission.ToString(), @"([A-Z][a-z]+)")
                                                    .Cast<Match>()
                                                    .Select(m => m.Value);

                                            var permissionWithSpaces = string.Join(" ", words);
                                            _userPermissions += permissionWithSpaces.ToString() + ", ";
                                        }
                                    }

                                }
                            }
                            catch(Exception ex)
                            {
                                message = "Could not load user permissions in web ." + ex.Message;
                            }
                        }
                        catch(Exception ex)
                        {
                            message = "Could find user in web ." + ex.Message;
                        }
                    }
                    catch(Exception ex)
                    {
                        message = "Could not connect to web ." + ex.Message;
                    }
                }
            }
            catch(Exception ex)
            {
                message = "Unexpected error occurred." + ex.Message;
            }

            return _userPermissions;
        }

        public bool IsTenantExternalSharingEnabled()
        {
            bool isTenantExternalShaingEnabled = false;
            string siteCollectionUrl = SPOAdminURL;
            string message = string.Empty;

            ClientContext clientContext;
            try
            {
                using (clientContext = GetClientContext(siteCollectionUrl, clientID, clientSecret))
                {
                    clientContext.ExecuteQuery();
                    Tenant currentO365Tenant = new Tenant(clientContext);
                    clientContext.Load(currentO365Tenant, O365t => O365t.SharingCapability);
                    clientContext.ExecuteQuery();
                    
                    SharingCapabilities _tenantSharing = currentO365Tenant.SharingCapability;
                    if (_tenantSharing == SharingCapabilities.Disabled)
                    {
                        message = "Sharing is currently disabled in our tenant.";
                        isTenantExternalShaingEnabled = false;
                    }
                    else
                        isTenantExternalShaingEnabled = true;
                }
            }
            catch(Exception ex)
            {
                message = ex.Message;
            }

           return isTenantExternalShaingEnabled;
        }

        public bool IsSiteCollectionExternalSharingEnabled(string _inputSiteCollectionUrl)
        {
            bool isSiteCollectionExternalShaingEnabled = false;
            string siteCollectionUrl = SPOAdminURL;
            string message = string.Empty;

            ClientContext clientContext;
            try
            {
                using (clientContext = GetClientContext(siteCollectionUrl, clientID, clientSecret))
                {
                    clientContext.ExecuteQuery();
                    Tenant currentO365Tenant = new Tenant(clientContext);
                    clientContext.Load(currentO365Tenant, O365t => O365t.SharingCapability);
                    clientContext.ExecuteQuery();
                   
                    SiteProperties _siteprops = currentO365Tenant.GetSitePropertiesByUrl(_inputSiteCollectionUrl, true);
                    clientContext.Load(_siteprops);
                    clientContext.ExecuteQuery();

                    var _currentShareSettings = _siteprops.SharingCapability;
                    if (_currentShareSettings == SharingCapabilities.Disabled)
                    {
                        isSiteCollectionExternalShaingEnabled = false;
                    }
                    else
                        isSiteCollectionExternalShaingEnabled = true;                       

                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }

            return isSiteCollectionExternalShaingEnabled;

        }

        public bool HasAccessGrantedToExternalUser(string _inputSiteCollectionUrl, string _strUserID)
        {
            bool isAccessGranted = false;
            string siteCollectionUrl = _inputSiteCollectionUrl;
            string message = string.Empty;

            ClientContext clientContext;
            try
            {
                using (clientContext = GetClientContext(siteCollectionUrl, clientID, clientSecret))
                {                   

                    var users = new List<UserRoleAssignment>();
                    users.Add(new UserRoleAssignment()
                    {
                        UserId = _strUserID,
                        Role = Role.Edit
                    });

                    WebSharingManager.UpdateWebSharingInformation(clientContext, clientContext.Web, users, true, "Access given by SOHA", true, true);
                    clientContext.ExecuteQuery();
                    isAccessGranted = true;

                    //string link = clientContext.Web.CreateAnonymousLinkForDocument("https://tenantname.sharepoint.com/Documents/sample.docx", ExternalSharingDocumentOption.View);
                    //SharingResult result = clientContext.Web.ShareDocument("https://tenantname.sharepoint.com/Documents/sample.docx", "someone@example.com", ExternalSharingDocumentOption.View, true, "Doc shared programmatically");
                }
            }
            catch(Exception ex)
            {
                isAccessGranted = false;
            }

            return isAccessGranted;
        }

        public List<string> GetSiteCollectionAdmins()
        {
            string message = string.Empty;
            string siteCollectionUrl = this.siteCollectionURL;
            List<string> lstSiteCollectionAdmins = new List<string>();
            lstSiteCollectionAdmins.Clear();

            try
            {
                if (string.IsNullOrEmpty(siteCollectionUrl))
                {
                    message = "No site collection URL is passed.";
                    
                }
                else
                {
                    Uri uriResult;
                    bool isUrlValid = Uri.TryCreate(siteCollectionUrl, UriKind.Absolute, out uriResult) && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
                    if (!isUrlValid)
                    {
                        message = "This is not a valid Uri.";
                    }
                    else
                    {
                        SecureString securePassword = new SecureString();
                        foreach (char c in SPOAdminPassowrd)
                        {
                            securePassword.AppendChar(c);
                        }

                        ClientContext clientContext;
                       
                        using (clientContext = GetClientContext(siteCollectionUrl, clientID, clientSecret))
                        {
                            
                            try
                            {
                                clientContext.ExecuteQuery();
                            }
                            catch (Exception ex)
                            {
                                message = ex.Message;
                            }

                            clientContext.Load(clientContext.Web);
                            clientContext.Load(clientContext.Site);
                            clientContext.Load(clientContext.Site.RootWeb);
                            clientContext.ExecuteQuery();

                            var users = clientContext.Site.RootWeb.SiteUsers;
                            clientContext.Load(users);
                            clientContext.ExecuteQuery();

                            foreach (var user in users)
                            {
                                if (user.IsSiteAdmin)
                                {  
                                    lstSiteCollectionAdmins.Add(user.Email.ToLower());                                    
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }
            return lstSiteCollectionAdmins;
        }

        public string GetSiteCollectionAdminsByPowerShell()
        {
            string siteCollectionAdmins = string.Empty;
            string siteCollectionUrl = this.siteCollectionURL;            

            RunspaceConfiguration runspaceConfiguration = RunspaceConfiguration.Create();

            Runspace runspace = RunspaceFactory.CreateRunspace(runspaceConfiguration);
            runspace.Open();

            RunspaceInvoke scriptInvoker = new RunspaceInvoke(runspace);
            scriptInvoker.Invoke("Set-ExecutionPolicy Unrestricted");

            Pipeline pipeline = runspace.CreatePipeline();

            string scriptfile = @"C:\LabaniPOCProjects\CollabLAMBot\PowerShell\test.ps1";          
            
            Command myCommand = new Command(scriptfile);
            CommandParameter siteCollectionParam = new CommandParameter("siteCollectionURL", siteCollectionUrl);
            myCommand.Parameters.Add(siteCollectionParam);
            CommandParameter SPOAdminParam = new CommandParameter("spoAdmin", SPOAdmin);
            myCommand.Parameters.Add(SPOAdminParam);
            CommandParameter SPOAdminPasswordParam = new CommandParameter("spoAdminPassword", SPOAdminPassowrd);
            myCommand.Parameters.Add(SPOAdminPasswordParam);
            CommandParameter SPOAdminURLParam = new CommandParameter("spoAdminURL", SPOAdminURL);
            myCommand.Parameters.Add(SPOAdminURLParam);


            pipeline.Commands.Add(myCommand);
            pipeline.Commands[0].MergeMyResults(PipelineResultTypes.Error, PipelineResultTypes.Output);

            //Execute PowerShell script
            Collection<PSObject> results = pipeline.Invoke();
            if(results.Count>=1)
            {                
                foreach(var result in results)
                {
                    siteCollectionAdmins += result.Properties["DisplayName"].Value.ToString() + " ;";
                }

            }            

            var error = pipeline.Error.ReadToEnd();
            runspace.Close();

            if (error.Count >= 1)
            {
                string errors = "";
                foreach (var Error in error)
                {
                    errors = errors + " " + Error.ToString();
                }
            }

            return siteCollectionAdmins;
        }
    }
}