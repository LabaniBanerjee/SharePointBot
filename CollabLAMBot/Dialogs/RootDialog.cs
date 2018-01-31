using Avanade.LAM.CollabBOT.LAM;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace CollabLAMBot.Dialogs
{
    [Serializable]
    public class RootDialog : IDialog<object>
    {
        public enum SupportType
        {
            UserAuthorization = 1,
            SiteCreation =2 ,
            SiteQuotaChange = 3 ,
            ExternalUserAccess =4 ,
            ProfileUpdates =5 
        }

        int _noDialog = -1;
        public SupportType? _spoSupportType;      


        public async Task StartAsync(IDialogContext context)
        {
            
            string currentLoggedinUser = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            context.UserData.SetValue<string>("Name", currentLoggedinUser);
            //await context.PostAsync($" {strGreetingContext} {currentLoggedinUser} , I am Collab Bot ."); 
            context.Wait(this.MessageReceivedAsync);
            //return Task.CompletedTask; // uncomment return if you remove async
        }


        public virtual async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> result)
        {
            var message = await result; // We've got a message!      // change this section with LUIS        

            if (message.Text.ToLower().Contains("authorization") )
            {
                _noDialog = Convert.ToInt32(SupportType.UserAuthorization);
            }
            else if(message.Text.ToLower().Contains("creation")
                || message.Text.ToLower().Contains("site"))
            {
                _noDialog = Convert.ToInt32(SupportType.SiteCreation);
            }
            else if (message.Text.ToLower().Contains("quota"))
            {
                _noDialog = Convert.ToInt32(SupportType.SiteQuotaChange);
            }
            else if (message.Text.ToLower().Contains("external"))
            {
                _noDialog = Convert.ToInt32(SupportType.ExternalUserAccess);
            }            
            else if (message.Text.ToLower().Contains("profile") 
                || message.Text.ToLower().Contains("update")
                || message.Text.ToLower().Contains("profile update"))
            {
                _noDialog = Convert.ToInt32(SupportType.ProfileUpdates);
            }
            else
            {
                //// User typed something else; for simplicity, ignore this input and wait for the next message.
                //context.Wait(this.MessageReceivedAsync);
                this.ShowOptions(context);
            }
            await this.SendWelcomeMessageAsync(context);
        }

        private void ShowOptions(IDialogContext context)
        {
            PromptDialog.Choice(context, this.OnOptionSelected, 
                new List<string>() { Constants.SiteAccess, Constants.SiteCreation, Constants.SiteQuotaChange, Constants.ExternalUserAccess, Constants.ProfileUpdates }, 
                "Are you looking for SharePoint Online Support?", 
                "Not a valid option",3);
        }

        private async Task OnOptionSelected(IDialogContext context, IAwaitable<string> result)
        {
            try
            {
                string optionSelected = await result;

                switch (optionSelected)
                {
                    case Constants.SiteAccess:
                        context.Call(new SiteAccessDialog(), this.ResumeAfterSiteAccessDialog);
                        break;

                    case Constants.SiteCreation:
                        context.Call(new SiteAccessDialog(), this.ResumeAfterSiteAccessDialog);
                        break;

                    case Constants.SiteQuotaChange:
                        context.Call(new SiteAccessDialog(), this.ResumeAfterSiteAccessDialog);
                        break;

                    case Constants.ExternalUserAccess:
                        context.Call(new SiteAccessDialog(), this.ResumeAfterSiteAccessDialog);
                        break;

                    case Constants.ProfileUpdates:
                        context.Call(new SiteAccessDialog(), this.ResumeAfterSiteAccessDialog);
                        break;
                }
            }
            catch (TooManyAttemptsException ex)
            {
                await context.PostAsync($"Ooops! Too many attemps :(. But don't worry, I'm handling that exception and you can try again!");

                context.Wait(this.MessageReceivedAsync);
            }
        }

        private async Task SendWelcomeMessageAsync(IDialogContext context)
        {
            await context.PostAsync("How can I help you?");
            if (_noDialog == 1)
            {
                context.Call(new SiteAccessDialog(), this.ResumeAfterSiteAccessDialog);
            }
            else if(_noDialog == 2)
            {
                await context.PostAsync("I'm sorry.I can't provide support for site creation.");
                context.Wait(this.MessageReceivedAsync);
            }
            else if (_noDialog == 3)
            {
                await context.PostAsync("I'm sorry.I can't provide support for site quota modification.");
                context.Wait(this.MessageReceivedAsync);
            }
            else if (_noDialog == 4)
            {
                await context.PostAsync("I'm sorry.I can't provide support for external user access.");
                context.Wait(this.MessageReceivedAsync);
            }
            else if (_noDialog == 5)
            {
                await context.PostAsync("I'm sorry.I can't provide support for user profile update.");
                context.Wait(this.MessageReceivedAsync);
            }
        }        


        private async Task ResumeAfterSiteAccessDialog(IDialogContext context, IAwaitable<string> result)
        {
            try
            {
                var resultFromSiteAccess = await result;

                await context.PostAsync($"Site Access dialog just told me this: {resultFromSiteAccess}");

                // Again, wait for the next message from the user.
                context.Wait(this.MessageReceivedAsync);
            }
            catch(TooManyAttemptsException)
            {
                await context.PostAsync("I'm sorry, I'm having issues understanding you. Let's try again.");

                await this.SendWelcomeMessageAsync(context);
            }
        }

    }
}