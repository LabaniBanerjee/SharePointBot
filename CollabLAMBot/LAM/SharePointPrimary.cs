using Avanade.LAM.CollabBOT.LAM;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
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

                    var newsite = new SiteCreationProperties()
                    {
                        Url = Constants.RootSiteCollectionURL+Constants.ManagedPath+_strSiteTitle,
                        Owner = _strPrimaryAdmin,
                        Template = "STS#0",
                        StorageMaximumLevel = 1,
                        UserCodeMaximumLevel = 1,
                        UserCodeWarningLevel = 1,
                        StorageWarningLevel = 1,
                        Title = _strSiteTitle+"_createdByBot",
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

        public Dictionary<string, string> SearchHelpArticleonTopic(int _inputHelpArticle)
        {
            string siteCollectionUrl = "https://avaindcollabsl.sharepoint.com";
            Dictionary<string, string> searchedDocs = new Dictionary<string, string>();

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
                                + "<FieldRef Name=\"Topic\"/>"                             
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

                    List<ListItem> currentSearchTopics = new List<ListItem>();
                    currentSearchTopics = lstListItemCollection.FindAll(a => Convert.ToString(a.FieldValuesAsHtml["Topic"]).Contains(_inputHelpArticle.ToString()));
                    if (currentSearchTopics != null && currentSearchTopics.Count > 0)
                    {
                        foreach (ListItem eachFoundArticle in currentSearchTopics)
                        {
                            searchedDocs.Add(eachFoundArticle.DisplayName, siteCollectionUrl+""+eachFoundArticle.File.ServerRelativeUrl);
                        }
                    }
                }
            }
            catch(Exception ex)
            {

            }
            return searchedDocs;

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

        public string GetSiteCollectionAdmins()
        {
            string siteCollectionAdmins = string.Empty;
            string siteCollectionUrl = this.siteCollectionURL;

            try
            {
                if (string.IsNullOrEmpty(siteCollectionUrl))
                {
                    return "No site collection URL is passed.";
                }
                else
                {
                    Uri uriResult;
                    bool isUrlValid = Uri.TryCreate(siteCollectionUrl, UriKind.Absolute, out uriResult) && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
                    if (!isUrlValid)
                    {
                        return "This is not a valid Uri.";
                    }
                    else
                    {
                        SecureString securePassword = new SecureString();
                        foreach (char c in SPOAdminPassowrd)
                        {
                            securePassword.AppendChar(c);
                        }

                        ClientContext clientContext;
                        //using (clientContext = new ClientContext(siteCollectionUrl))
                        using (clientContext = GetClientContext(siteCollectionUrl, clientID, clientSecret))
                        {
                            //clientContext.Credentials = new SharePointOnlineCredentials(SPOAdmin, securePassword);
                            try
                            {
                                clientContext.ExecuteQuery();
                            }
                            catch (Exception ex)
                            {
                                return "Error : " + ex.Message;
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
                                    siteCollectionAdmins += user.Email + "\r\r";
                                }

                            }
                        }//
                    }
                }
            }//
            catch (Exception ex)
            {
                return "Error : "+ex.Message;
            }
            return siteCollectionAdmins;
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