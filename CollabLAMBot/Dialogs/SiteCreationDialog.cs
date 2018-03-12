using Avanade.LAM.CollabBOT.LAM;
using CollabLAMBot.LAM;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.FormFlow;
using Microsoft.Bot.Connector;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
namespace CollabLAMBot.Dialogs
{
    [Serializable]
    public class SiteCreationDialog : IDialog<string>
    {
        public async Task StartAsync(IDialogContext context)
        {
            await context.PostAsync("Sure, I can help you with Site Creation, before that I need 2 inputs from you.");

           // await context.PostAsync("Sure, I can help you with Site Creation, before that I need 2 inputs from you.[hello](https://www.microsoft.com)");

            var SiteCreationFormDialog = FormDialog.FromForm(this.BuildSiteCreationForm, FormOptions.PromptInStart);

            context.Call(SiteCreationFormDialog, this.ResumeSiteCreationDialog);            
        }

        private async Task ResumeSiteCreationDialog(IDialogContext context, IAwaitable<SiteCreationQuery> result)
        {

            var resultFromSiteCreation = await result;
            string _strSiteTitle = resultFromSiteCreation.SiteCollectionTitle;
            string _strPrimaryAdmin = resultFromSiteCreation.PrimarySiteCollectionAdmin;
            SharePointPrimary obj = new SharePointPrimary();
            try
            {
                if(!string.IsNullOrEmpty(_strSiteTitle)&& !string.IsNullOrEmpty(_strPrimaryAdmin))
                {
                    try
                    {
                        
                        Task<bool> _isCreated = obj.IsSiteCollectionCreated(_strSiteTitle, _strPrimaryAdmin);                        

                        Attachment attachment = new Attachment();
                        attachment.ContentType = "application/pdf";
                        attachment.ContentUrl = Constants.RootSiteCollectionURL + "" + Constants.ManagedPath + "" + _strSiteTitle;
                        attachment.Name = _strSiteTitle;

                        var replyMessage = context.MakeMessage();
                        replyMessage.Attachments = new List<Attachment> { attachment };

                        await context.PostAsync("I have started creating your site collection. It may take 5 - 10 minutes to complete the process."+
                            "\r\r Please browse the site ['SOHA_" + _strSiteTitle + "'](" + Constants.RootSiteCollectionURL + "" + Constants.ManagedPath + "" + _strSiteTitle + ") after sometime.");
                        //await context.PostAsync("Please browse the site '" + Constants.RootSiteCollectionURL + "" + Constants.ManagedPath + "" + _strSiteTitle + "' after sometime.");
                        context.Done("Done");

                    }
                    catch (Exception)
                    {
                        context.Fail(new Exception("Unable to create site collection."));
                        await context.PostAsync("Sorry \U0001F641 , I am unable to create site collection for you. Let us try again.");
                    }                    
                }

            }
            catch(TooManyAttemptsException)
            {
                await context.PostAsync("Sorry \U0001F641 , I am unable to understand you. Let us try again.");
            }
        }

        private IForm<SiteCreationQuery> BuildSiteCreationForm()
        {
            OnCompletionAsyncDelegate<SiteCreationQuery> processUserAccessSearch = async (context, authorizationState) =>   {
                //await context.PostAsync($"Okay. Creating site collection ...");
                await context.PostAsync($"Thank you!");
            };

            return new FormBuilder<SiteCreationQuery>()
                .Field(nameof(SiteCreationQuery.SiteCollectionTitle),validate: ValidateSiteCollectionURL)
                .Field(nameof(SiteCreationQuery.PrimarySiteCollectionAdmin), validate: ValidatePrimaryAdmin)               
                .AddRemainingFields()
                .Confirm("Great. I am ready to submit your request with the following details \U0001F447 " +
                        "\r\r The new site collection URL will be '" + Constants.RootSiteCollectionURL+ ""+Constants.ManagedPath+"SOHA_{SiteCollectionTitle} '"+
                        "\r\r and the primary administartor {PrimarySiteCollectionAdmin}. " +
                        "\r\r Is that correct?")
                .OnCompletion(processUserAccessSearch)
                 //.Message("Thank you, I have submitted your request.")
                .Build();

        }

        private async Task<ValidateResult> ValidateSiteCollectionURL(SiteCreationQuery state, object value)
        {
            string _inputSiteCollectionTitle = Convert.ToString(value);
            var result = new ValidateResult { IsValid = false, Value = _inputSiteCollectionTitle };

            SharePointPrimary obj = new SharePointPrimary();
            result.IsValid = obj.DoesContainSpecialCharacter(_inputSiteCollectionTitle);
            
            if (result.IsValid)
            {
                result.IsValid = !obj.DoesURLExist(_inputSiteCollectionTitle);
                if (!result.IsValid)
                    result.Feedback = $"A site collection with url '{_inputSiteCollectionTitle}' is already present in our tenant. Try something different.";
                else
                    result.Feedback = "The site title is available to use.";
            }
            else
            {
               result.Feedback = "Special characters are not allowed in site collection URL.";
            }

            return result;
        }

        private async Task<ValidateResult> ValidatePrimaryAdmin(SiteCreationQuery state, object value)
        {
            string _inputSPOUserID = Convert.ToString(value);
            var result = new ValidateResult { IsValid = false, Value = _inputSPOUserID };

            SharePointPrimary obj = new SharePointPrimary();
            result.IsValid = obj.IsValidSPOUser(_inputSPOUserID);
            if (!result.IsValid)
                result.Feedback = $"I could not find the site collection administrator's profile {_inputSPOUserID} in our O365 tenant. ";
            else
            {
                result.Feedback = "Yes. I have found the primary site collection administrator's profile in our O365 tenant. ";                                      
               
            }
            return result;
        }
    }

    #region query helper

    [Serializable]
    public class SiteCreationQuery
    {
        [Prompt("Please enter a site collection title (avoid any special characters like $,#,!).")]        
        //[Pattern(@"(<Undefined control sequence>\d)^[a-zA-Z0-9 _]*$")]
        public string SiteCollectionTitle { get; set; }

        [Prompt("Please enter Primary Site Collection Admin id.")]
        public string PrimarySiteCollectionAdmin { get; set; }        
    }

    #endregion

}