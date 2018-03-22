using CollabLAMBot.LAM;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.FormFlow;
using System;
using System.Threading.Tasks;


namespace CollabLAMBot.Dialogs
{
    [Serializable]
    public class ProfileUpdatesDialog : IDialog
    {
        public async Task StartAsync(IDialogContext context)
        {
            await context.PostAsync("Sure, I can update your profile, before that I need 2 inputs from you. ");

            var userProfileFormDialog = FormDialog.FromForm(this.BuildUserProfileForm, FormOptions.PromptInStart);

            context.Call(userProfileFormDialog, this.ResumeUserProfileDialog);            
        }

        private async Task ResumeUserProfileDialog(IDialogContext context, IAwaitable<ProfileUpdateQuery> result)
        {
            try
            {
                var resultFromUserProfile = await result;
                
                string _strUserID = resultFromUserProfile.SPOUserID;
                string _strProperty = Convert.ToString(resultFromUserProfile.UserProfilePropertyToUpdate);
                bool _isProfileUpdated = false;

                SharePointPrimary obj = new SharePointPrimary();
                
                try
                {
                    _isProfileUpdated = obj.IsUserProfilePropertyUpdated( _strUserID, _strProperty);

                    if (_isProfileUpdated)
                    {
                        await context.PostAsync($"Profile Updated \U00002705  ");
                        context.Done("Done");
                    }
                    else
                    {
                        await context.PostAsync("Profile could not be updated \U0001F641 Please try again later.");
                        context.Done("Not Done");
                    }
                }
                catch (Exception ex)
                {
                    context.Fail(new TooManyAttemptsException("Unable to user profile \U0001F641 Please try again later."));
                }
                
            }
            catch (TooManyAttemptsException)
            {
                await context.PostAsync("Sorry \U0001F641 I am unable to understand you. Let us try again.");
            }
        }



        private IForm<ProfileUpdateQuery> BuildUserProfileForm()
        {
            OnCompletionAsyncDelegate<ProfileUpdateQuery> processUserProfileUpdate = async (context, authorizationState) =>
            {
                await context.PostAsync($"Processing your request...");
            };

            return new FormBuilder<ProfileUpdateQuery>()                
                .Field(nameof(ProfileUpdateQuery.SPOUserID), validate: ValidateSPOUserID)
                .Field(nameof(ProfileUpdateQuery.UserProfilePropertyToUpdate))
                .AddRemainingFields()
                .Confirm("Great. I am ready to submit your request with the following details \U0001F447 " +
                        "\r\r Your id is {SPOUserID} and you want to update '{UserProfilePropertyToUpdate}' in your profile. " +
                        "\r\r Is that correct?")
                .OnCompletion(processUserProfileUpdate)
                .Message("Thank you!")
                .Build();
        }


        private async Task<ValidateResult> ValidateSPOUserID(ProfileUpdateQuery state, object value)
        {
            string _inputSPOUserID = Convert.ToString(value);
            var result = new ValidateResult { IsValid = false, Value = _inputSPOUserID };

            SharePointPrimary obj = new SharePointPrimary();
            result.IsValid = obj.IsValidSPOUser(_inputSPOUserID);
            if (!result.IsValid)
                result.Feedback = $"I could not find your profile {_inputSPOUserID} in our O365 tenant. ";
            else
            {              
                result.Feedback = "Yes. I have found your profile in our O365 tenant. ";
            }

            return result;
        }

    }

    #region query helper

    [Serializable]
    public class ProfileUpdateQuery
    {        
        [Prompt("Could you please help me with your SharePoint Online user id?")]
        public string SPOUserID { get; set; }
        
        public ProfileProperties? UserProfilePropertyToUpdate { get; set; }

        public enum ProfileProperties
        {
            PreferredName = 2,
            AccountName = 3,           
            JobTitle = 4,
            WorkPhone = 6
        }
    }

    #endregion
}