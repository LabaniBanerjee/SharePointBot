using CollabLAMBot.LAM;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.FormFlow;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;


namespace CollabLAMBot.Dialogs
{
    [Serializable]
    public class ExternalUserAccessDialog : IDialog
    {
        public async Task StartAsync(IDialogContext context)
        {
            await context.PostAsync("Sure, I can help ypu with External Sharing, before that I need 2 inputs from you.");


            var externalSharingFormDialog = FormDialog.FromForm(this.BuildExternalSharingForm, FormOptions.PromptInStart);

            context.Call(externalSharingFormDialog, this.ResumeExternalSharingDialog);

        }

        private async Task ResumeExternalSharingDialog(IDialogContext context, IAwaitable<ExternalSharingQuery> result)
        {
            try
            {
                var resultFromExtrnalSharing = await result;
                string _strURL = resultFromExtrnalSharing.SiteCollectionURL;
                string _strUserID = resultFromExtrnalSharing.SPOExternalUserID;
                bool _isPermissionGranted = false;
                bool _isSiteCollectionSharingEnabled = false;
                bool _isTenantSharingEnabled = false;

                try
                {
                    SharePointPrimary obj = new SharePointPrimary(_strURL);

                    List<string> lstSiteCollectionAdmins = new List<string>();
                    lstSiteCollectionAdmins = obj.GetSiteCollectionAdmins();
                    string strSiteCollectionAdmins = string.Empty;

                    foreach (var eachAdmin in lstSiteCollectionAdmins)
                    {
                        strSiteCollectionAdmins += eachAdmin + "; ";
                    }

                    _isTenantSharingEnabled = obj.IsTenantExternalSharingEnabled();
                    if(_isTenantSharingEnabled)
                    {
                        _isSiteCollectionSharingEnabled = obj.IsSiteCollectionExternalSharingEnabled(_strURL);
                        if (_isSiteCollectionSharingEnabled)
                        {
                            _isPermissionGranted = obj.HasAccessGrantedToExternalUser(_strURL, _strUserID);
                            
                            if (_isPermissionGranted)
                            {
                                await context.PostAsync($"Access granted \U00002705 An email is sent to '{_strUserID}' ");
                                context.Done("External Access granted.");
                            }
                            else
                            {
                                await context.PostAsync("Permission could not be granted \U0001F641 Please try again later.");
                                context.Done("External Access could not be granted.");
                            }
                        }
                        else
                        {                            
                            await context.PostAsync($"Access could not be granted \U0001F641 External Sharing is disabled for this Site Collection.");
                            await context.PostAsync($"Please reach out to one of the Site Collection Adminstrators listed below:" +
                               "\r\r" + strSiteCollectionAdmins);
                            context.Done("Access could not be granted. External Sharing is disabled for this Site Collection.");
                        }
                    }
                    else
                    {                        
                        await context.PostAsync($"Access could not be granted \U00002705 External Sharing is disabled at our Tenant level. ");
                        await context.PostAsync($"Please reach out to one of the Site Collection Adminstrators listed below:" +
                               "\r\r" + strSiteCollectionAdmins);
                        context.Done("Access could not be granted. External Sharing is disabled at our Tenant level.");
                    }                   
                }
                catch(Exception)
                {
                    context.Fail(new Exception("Unable to grant permission to user \U0001F641 . Please try again later."));
                }
            }
            catch (TooManyAttemptsException)
            {
                await context.PostAsync("Sorry \U0001F641 , I am unable to understand you. Let us try again.");
            }
        }


        private IForm<ExternalSharingQuery> BuildExternalSharingForm()
        {
            OnCompletionAsyncDelegate<ExternalSharingQuery> processExternalSharing = async (context, authorizationState) =>
            {
                await context.PostAsync($"Processing your request...");

            };

            return new FormBuilder<ExternalSharingQuery>()
                .Field(nameof(ExternalSharingQuery.SiteCollectionURL), validate: ValidateSiteCollectionURL)
                .Field(nameof(ExternalSharingQuery.SPOExternalUserID))               
                .AddRemainingFields()
                .Confirm("Great. I am ready to submit your request with the following details \U0001F447 " +
                        "\r\r Your id is {SPOExternalUserID} and you want to access Site Collection '{SiteCollectionURL}'. " +
                        "\r\r Is that correct?")
                .OnCompletion(processExternalSharing)
                .Message("Thank you! I have submitted your request.")
                .Build();
        }

        private async Task<ValidateResult> ValidateSiteCollectionURL(ExternalSharingQuery state, object value)
        {
            string _inputSiteCollectionURL = Convert.ToString(value);
            var result = new ValidateResult { IsValid = false, Value = _inputSiteCollectionURL };

            SharePointPrimary obj = new SharePointPrimary(_inputSiteCollectionURL);
            result.IsValid = obj.IsValidSiteCollectionURL();
            if (!result.IsValid)
                result.Feedback = $"This site {_inputSiteCollectionURL} does not exist in our O365 tenant. Please enter a valid URL.";
            else
                result.Feedback = "Yes. I have found this site collection in our O365 tenant.";

            return result;
        }




        #region query helper

        [Serializable]
        public class ExternalSharingQuery
        {
            [Prompt("Please enter the Site Collection URL.")]
            public string SiteCollectionURL { get; set; }

            [Prompt("Could you please help me with your id?")]
            public string SPOExternalUserID { get; set; }

            
        }

        #endregion


    }
}