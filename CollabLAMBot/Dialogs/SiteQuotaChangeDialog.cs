using CollabLAMBot.LAM;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.FormFlow;
using Microsoft.Bot.Builder.FormFlow.Advanced;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace CollabLAMBot.Dialogs
{
    [Serializable]
    public class SiteQuotaChangeDialog : IDialog
    {
        public async Task StartAsync(IDialogContext context)
        {            
            await context.PostAsync("Sure, I can update the Site Collection Storage Quota, before that I need 2 inputs from you.");            

            var siteQuotaFormDialog = FormDialog.FromForm(this.BuildSiteQuotaForm, FormOptions.PromptInStart);

            context.Call(siteQuotaFormDialog, this.ResumeSiteQuotaDialog);
        }

        private async Task ResumeSiteQuotaDialog(IDialogContext context, IAwaitable<SiteQuotaChangeQuery> result)
        {
            try
            {
                var resultFromSiteQuotaChange = await result;
                string _strURL = resultFromSiteQuotaChange.SiteCollectionURL;
                string _strUserID = resultFromSiteQuotaChange.SPOUserID;
                //int _intNewQuota = Convert.ToInt32(resultFromSiteQuotaChange.NewStorageQuota);
                int _intNewQuota = 1;

                if (!string.IsNullOrEmpty(_strURL) && !string.IsNullOrEmpty(_strUserID) && (_intNewQuota != 0))
                {
                    try
                    {
                        SharePointPrimary obj = new SharePointPrimary(_strURL);
                        bool isSiteCollectionAdmin = obj.IsSiteCollectionAdmin(_strUserID);
                        if (isSiteCollectionAdmin)
                        {
                            bool isStorageUpdated = obj.IsSiteCollectionStorageQuotaUpdated(_strURL, _intNewQuota);
                            context.Done("Storage Quota Updated.");
                            if (isStorageUpdated)
                            {
                                await context.PostAsync("The Storage Quota for the Site Collection is being updated. Please refresh the site after sometime.");
                            }
                            else
                                await context.PostAsync("Storage quota could not be updated \U0001F641 . Please try again later.");
                        }
                        else
                        {
                            List<string> lstSiteCollectionAdmins = new List<string>();
                            lstSiteCollectionAdmins = obj.GetSiteCollectionAdmins();
                            string strSiteCollectionAdmins = string.Empty;

                            foreach(var eachAdmin in lstSiteCollectionAdmins)
                            {
                                strSiteCollectionAdmins += eachAdmin+ ";";
                            }

                            await context.PostAsync($"Sorry \U0001F641 I just found out that you are not authorized to update the Storage Quota of the Site Collection '{_strURL}'");
                            await context.PostAsync($"Please reach out to one of the Site Collection Adminstrators listed below:" +
                                "\r\r"+ strSiteCollectionAdmins);

                            context.Done("Storage Quota Not Updated.");
                        }
                    }
                    catch(Exception)
                    {
                        context.Fail(new TooManyAttemptsException("Unable to update Storage Quota \U0001F641 . Please try again later."));
                    }
                }
            }
            catch (TooManyAttemptsException)
            {
                await context.PostAsync("Sorry \U0001F641 , I am unable to understand you. Let us try again.");
            }
        }



        private IForm<SiteQuotaChangeQuery> BuildSiteQuotaForm()
        {
            OnCompletionAsyncDelegate<SiteQuotaChangeQuery> processStorageQuota = async (context, authorizationState) =>
            {
                await context.PostAsync($"Processing your request...");
               
            };

            return new FormBuilder<SiteQuotaChangeQuery>()
                .Field(nameof(SiteQuotaChangeQuery.SiteCollectionURL), validate: ValidateSiteCollectionURL)
                .Field(nameof(SiteQuotaChangeQuery.SPOUserID), validate: ValidateSPOUserID)                               
                //.Field(nameof(SiteQuotaChangeQuery.NewStorageQuota))               
                .AddRemainingFields()
                .Confirm("Great. I am ready to submit your request with the following details \U0001F447 " +
                        "\r\r Your id is {SPOUserID} and you want to update storage quota for the Site Collection '{SiteCollectionURL}'. " +
                        "\r\r Is that correct?")
                .OnCompletion(processStorageQuota)
                .Message("Thank you! I have submitted your request.")
                .Build();
        }

        private async Task<ValidateResult> ValidateSiteCollectionURL(SiteQuotaChangeQuery state, object value)
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

        private async Task<ValidateResult> ValidateSPOUserID(SiteQuotaChangeQuery state, object value)
        {

            string _inputSPOUserID = Convert.ToString(value);
            var result = new ValidateResult { IsValid = false, Value = _inputSPOUserID };

            SharePointPrimary obj = new SharePointPrimary(state.SiteCollectionURL);
            result.IsValid = obj.IsValidSPOUser(_inputSPOUserID);
            if (!result.IsValid)
                result.Feedback = $"I could not find the profile {_inputSPOUserID} in our O365 tenant. ";
            else            
                result.Feedback = "Yes, I have found your profile in our O365 tenant.";
                               
            return result;
        }




        #region query helper

        [Serializable]
        public class SiteQuotaChangeQuery
        {
            [Prompt("Please enter the Site Collection URL for which you want to update the Storage Quota.")]
            public string SiteCollectionURL { get; set; }

            [Prompt("Could you please help me with your SPO user id?")]
            public string SPOUserID { get; set; }

            //[Prompt("What is the maximum size (in GB) you want to assign to this Site Collection? (Enter integer value 2 to 10)")]
            //[Numeric(2, 10)]
            //public int NewStorageQuota { get; set; }
        }

        #endregion



    }
}