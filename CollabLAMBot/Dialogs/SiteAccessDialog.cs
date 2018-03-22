using Avanade.LAM.CollabBOT.LAM;
using CollabLAMBot.LAM;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.FormFlow;
using Microsoft.Bot.Connector;
using System;
using System.Threading.Tasks;

namespace CollabLAMBot.Dialogs
{
    [Serializable]
    public class SiteAccessDialog : IDialog<string>
    {
        private int attempts = 3;
        public string siteURL = string.Empty;

        #region without Form Single quesion Implementation 
        //public async Task StartAsync(IDialogContext context)
        //{
        //    await context.PostAsync("What is site collection URL?");

        //    context.Wait(this.MessageReceivedAsync);
        //}

        //private async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> result)
        //{
        //    var message = await result;

        //    if ((message.Text != null) && (message.Text.Trim().Length > 0))
        //    {
        //        SharePointPrimary obj = new SharePointPrimary(message.Text);
        //        if (obj.IsValidSiteCollectionURL())
        //        {
        //            string siteCollectionAdmins = obj.GetSiteCollectionAdmins();
        //            //context.Done(message.Text);
        //            context.Done("Site collection admin(s) are : " + siteCollectionAdmins);
        //        }
        //        else
        //        {
        //            context.Fail(new TooManyAttemptsException("The URL was not a valid site collection URL."));
        //        }

        //    }
        //    else
        //    {
        //        --attempts;
        //        if (attempts > 0)
        //        {
        //            await context.PostAsync("I'm sorry, I don't understand your reply. What is URL (e.g. 'https://avaindcollabsl.sharepoint.com')?");

        //            context.Wait(this.MessageReceivedAsync);
        //        }
        //        else
        //        {
        //            context.Fail(new TooManyAttemptsException("Message was not a string or was an empty string."));
        //        }
        //    }
        //}

        #endregion


        public async Task StartAsync(IDialogContext context)
        {
            await context.PostAsync("Sure, I can grant you access to a site, before that I need 3 inputs from you.");

            var userAuthorizationFormDialog = FormDialog.FromForm(this.BuildUserAuthorizationForm, FormOptions.PromptInStart);

            context.Call(userAuthorizationFormDialog, this.ResumeUserAuthorizationDialog);
        }        

        private async Task ResumeUserAuthorizationDialog(IDialogContext context, IAwaitable<UserAuthorizationQuery> result)
        {
            try
            {
                var resultFromUserAuthorization = await result;
                string _strURL = resultFromUserAuthorization.SiteCollectionURL;
                string _strUserID = resultFromUserAuthorization.SPOUserID;
                int _intRoleype = Convert.ToInt32(resultFromUserAuthorization.SharepointOnlineRole);               
                bool _isPermissionGranted = false;

                SharePointPrimary obj = new SharePointPrimary();    
                           

                if (!string.IsNullOrEmpty(_strURL) && !string.IsNullOrEmpty(_strUserID) && (_intRoleype != 255))
                {
                    try
                    {
                        _isPermissionGranted = obj.HasPermissionGrantedToUser(_strURL, _strUserID, _intRoleype);
                        
                        if (_isPermissionGranted)
                        {
                            await context.PostAsync($"Access granted \U00002705 Please browse the site url '{_strURL}' ");
                            context.Done("Done");
                        }
                        else
                        {
                            await context.PostAsync("Permission could not be granted \U0001F641 . Please try again later.");
                            context.Done("Not Done");
                        }

                    }
                    catch(Exception ex)
                    {
                        context.Fail(new TooManyAttemptsException("Unable to grant permission to user \U0001F641 . Please try again later."));
                    }
                }
            }
            catch (TooManyAttemptsException)
            {
                await context.PostAsync("Sorry \U0001F641 , I am unable to understand you. Let us try again.");               
            }
        }

        private IForm<UserAuthorizationQuery> BuildUserAuthorizationForm()
        {
            OnCompletionAsyncDelegate<UserAuthorizationQuery> processUserAccessSearch = async (context, authorizationState) =>
            {
                await context.PostAsync($"Processing your access...");
            };

            return new FormBuilder<UserAuthorizationQuery>()
                .Field(nameof(UserAuthorizationQuery.SiteCollectionURL),validate : ValidateSiteCollectionURL)                
                .Field(nameof(UserAuthorizationQuery.SPOUserID),validate : ValidateSPOUserID)
                .Field(nameof(UserAuthorizationQuery.SharepointOnlineRole))
                .Message("If you want to know more about permissions, please visit : " + Constants.RoleTypeMSDN)
                .AddRemainingFields()
                .Confirm("Great. I am ready to submit your request with the following details \U0001F447 " +
                        "\r\r Your id is {SPOUserID} and you need '{SharepointOnlineRole}' access to site {SiteCollectionURL}. " +
                        "\r\r Is that correct?")                 
                .OnCompletion(processUserAccessSearch)
                .Message("Thank you! I have submitted your request.")
                .Build();
        }

        private async Task<ValidateResult> ValidateSPOUserID(UserAuthorizationQuery state, object value)
        {

            string _inputSPOUserID = Convert.ToString(value);
            var result = new ValidateResult { IsValid = false, Value = _inputSPOUserID };

            SharePointPrimary obj = new SharePointPrimary(state.SiteCollectionURL);
            result.IsValid = obj.IsValidSPOUser(_inputSPOUserID);
            if (!result.IsValid)
                result.Feedback = $"I could not find your profile {_inputSPOUserID} in our O365 tenant. " +
                                    "\r\r If you are an external user, please raise your request in the 'External User Access' option" +
                                    "\r\r or enter a valid user id.";
            else
            {
                // call method to check permission 
                string _userPermission = obj.CheckUserPermission(state.SiteCollectionURL, _inputSPOUserID);
                if (string.IsNullOrEmpty(_userPermission))
                    result.Feedback = "Yes. I have found your profile in our O365 tenant. ";// +
                                                                                            //"\r\r To know more about permission, please visit : "+ Constants.RoleTypeMSDN;
                else
                    result.Feedback = $"Looks like you already have some permissions assigned to you : " +
                                        $"\r\r {_userPermission}";// +
                                       // "\r\r To know more about permission,please visit : " + Constants.RoleTypeMSDN;
            }

            return result;
        }

        private async Task<ValidateResult> ValidateSiteCollectionURL(UserAuthorizationQuery state, object value)
        {
            string _inputSiteCollectionURL = Convert.ToString(value);
            var result = new ValidateResult { IsValid = false, Value = _inputSiteCollectionURL };

            SharePointPrimary obj = new SharePointPrimary(_inputSiteCollectionURL);
            result.IsValid = obj.IsValidSiteCollectionURL();
            if (!result.IsValid)
                result.Feedback = $"This site {_inputSiteCollectionURL} does not exist in our O365 tenant. Please enter a valid URL.";
            else
                result.Feedback = "Yes. I have found this site in our O365 tenant.";

            return result;
        }
    }

    #region query helper

    [Serializable]
    public /*internal */class UserAuthorizationQuery
    {
        [Prompt("Please share the Site Collection URL you want to access.")]
        public string SiteCollectionURL { get; set; }

        [Prompt("Could you please help me with your SharePoint Online user id?")]
        public string SPOUserID { get; set; }

        //[Template(TemplateUsage.EnumSelectOne, "What is the role type you look for?)" )]
        //[Prompt("What is the role type you look for?")]
        public CustomRoleType? SharepointOnlineRole { get; set; }

        public enum CustomRoleType
        {
            //None = 0,
            // Guest = 1, /*cannot be assigned programatically */
           
            Reader = 2,
            Contributor = 3,
            [Describe("Web Designer")]
            WebDesigner = 4,
            [Describe("Administrator")]
            Administrator = 5,
            Editor = 6
        }
    }

    #endregion
}