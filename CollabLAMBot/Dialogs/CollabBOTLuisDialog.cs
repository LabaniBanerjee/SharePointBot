using Avanade.LAM.CollabBOT.Dialogs;
using Avanade.LAM.CollabBOT.LAM;
using Avanade.LAM.CollabBOT.Utility;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.FormFlow;
using Microsoft.Bot.Builder.Internals.Fibers;
using Microsoft.Bot.Builder.Luis;
using Microsoft.Bot.Builder.Luis.Models;
using Microsoft.Bot.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CollabLAMBot.Dialogs
{
    [LuisModel("0813b523-b87f-451c-b8d2-3c6224154886", "6f3f8ca7506a48eda1778f0758a27190" ,domain: "southeastasia.api.cognitive.microsoft.com")]
    [Serializable]
    public class CollabBOTLuisDialog : LuisDialog<object>
    {
               

        #region constructor
        public CollabBOTLuisDialog(params ILuisService[] services) : base(services) { }

        #endregion      


        #region intents

        [LuisIntent("Help")]
        [LuisIntent("None")]
        [LuisIntent("")]
        public async Task None(IDialogContext context, LuisResult result)
        {
            #region commented prompt section
            //var _helpdialog = new PromptDialog.PromptChoice<string>(
            //    new string[] {
            //        Constants.SiteAccess,
            //        Constants.SiteCreation,
            //        Constants.SiteQuotaChange,
            //        Constants.ExternalUserAccess,
            //        Constants.ProfileUpdates,
            //        Constants.Help},
            //    "May I assist you with something else?",
            //    "Please select between 1 to 6.",
            //    3,
            //    PromptStyle.PerLine,
            //    new string[] {
            //        "Site Access (press 1)",
            //        "Site/ Site Collection Creation (press 2)",
            //        "Update Site Collection Quota (press 3)",
            //        "Give access to external user (press 4)",
            //        "Update user profile (press 5)",
            //         "Help (press 6)"});            

            //context.Call(_helpdialog, OnOptionSelected);
            #endregion                      

            await context.PostAsync("Sorry \U0001F641, I am unable to understand you." +                
                        "\r\r Do you want me to raise a service request for the same?");

            context.Call(new ServiceNowDialog(), Callback);
        }        

        [LuisIntent("ArticleSearch")]
        public async Task ArticleSearch(IDialogContext context, LuisResult result)
        {
            IntentRecommendation re = result.TopScoringIntent;
            string test = re.Intent;
            //await context.PostAsync(test);
            context.Call(new ArticleSearch(test) ,Callback);
        }

        [LuisIntent("Thanks")]
        public async Task ThankGreetingCall(IDialogContext context, LuisResult result)
        { context.Call(new ThankGreetingDialog(), Callback); }

        [LuisIntent("Greeting")]       
        public async Task WelcomeGreetingCall(IDialogContext context, LuisResult result)
        {   context.Call(new GreetingDialog2(), Callback); }

        [LuisIntent("Bye")]
        public async Task FarewellGreetingCall(IDialogContext context, LuisResult result)
        {   context.Call(new FarewellGreeting(), Callback);}        

        [LuisIntent("UserAuthorization")]
        public async Task UserAuthorization(IDialogContext context, LuisResult result)
        {   context.Call(new SiteAccessDialog(), Callback);}

        [LuisIntent("SiteCreation")]
        public async Task SiteCreationModule(IDialogContext context, LuisResult result)
        {   context.Call(new SiteCreationDialog(), Callback);}

        [LuisIntent("SiteQuotaChange")]
        public async Task SiteQuotaChangeModule(IDialogContext context, LuisResult result)
        {   context.Call(new SiteQuotaChangeDialog(), Callback);}

        [LuisIntent("ExternalUserAccess")]
        public async Task ExternalUserAccessModule(IDialogContext context, LuisResult result)
        {   context.Call(new ExternalUserAccessDialog(), Callback);}

        [LuisIntent("ProfileUpdates")]
        public async Task ProfileUpdatesModule(IDialogContext context, LuisResult result)
        {
            IntentRecommendation re = result.TopScoringIntent;
            string test = re.Intent;
            
            context.Call(new ProfileUpdatesDialog(), Callback);}

        #endregion

        #region private methods and callback

        /// <summary>
        /// We decided to remove prompt functionality
        /// </summary>
        /// <param name="context"></param>
        private void ShowOptions(IDialogContext context)
        {
            PromptDialog.Choice(context:context,resume:this.OnOptionSelected,
                options:
                new string[] {
                    Constants.SiteAccess,
                    Constants.SiteCreation,
                    Constants.SiteQuotaChange,
                    Constants.ExternalUserAccess,
                    Constants.ProfileUpdates,
                    Constants.Help},
                prompt:"These are few SharePoint Online support options : ",
                retry:"Please select between 1 to 6.",
                attempts:3,
                promptStyle:PromptStyle.PerLine,descriptions:
                new string[] {
                    "Site Access (press 1)",
                    "Site/ Site Collection Creation (press 2)",
                    "Update Site Collection Quota (press 3)",
                    "Give access to external user (press 4)",
                    "Update user profile (press 5)",
                     "Help (press 6)"});

        }

        /// <summary>
        /// We decided to remove prompt functionality
        /// </summary>
        private async Task OnOptionSelected(IDialogContext context, IAwaitable<string> result)
        {
            try
            {
                string optionSelected = await result;
                await context.PostAsync("You selected : " + optionSelected);

                switch (optionSelected)
                {
                    case Constants.SiteAccess:
                        context.Call(new SiteAccessDialog(), this.Callback);
                        break;

                    case Constants.SiteCreation:
                        context.Call(new SiteCreationDialog(), this.Callback);
                        break;

                    case Constants.SiteQuotaChange:
                        context.Call(new SiteQuotaChangeDialog(), this.Callback);
                        break;

                    case Constants.ExternalUserAccess:
                        context.Call(new ExternalUserAccessDialog(), this.Callback);
                        break;

                    case Constants.ProfileUpdates:
                        context.Call(new ProfileUpdatesDialog(), this.Callback);
                        break;

                    case Constants.Help:
                        context.Call(new ArticleSearch(), this.Callback);
                        break;
                }
            }
            catch (TooManyAttemptsException ex)
            {
                await context.PostAsync($"Ooops! Too many attemps :(. But don't worry, I'm handling that exception and you can try again!");
                context.Wait(MessageReceived);
            }
        }                       
        private async Task Callback(IDialogContext context, IAwaitable<object> result)
        {
            var options = "* Site Access \n" +
                            "* Site Creation \n" +
                            "* Site Quota Change \n" +
                            "* External User Access \n" +
                            "* Profile Updates \n" +
                            "* User Guides/ KB Articles \n";

            try
            {
                var myresult = await result;
                
                if (!(myresult.Equals("Greeting") || myresult.Equals("service now exited with no") || myresult.Equals("Farewell")))
                {
                    await context.PostAsync("Hope I am able to assist you on your request.");

                    await context.PostAsync("However, I can help you with some other popular Support Requests too such as: " +
                     "\r\r" + options +
                            "\r\r Please type any of the above keywords or any other query you may have.");
                }

            }
            catch (FormCanceledException)
            {
                await context.PostAsync("Looks like you wanted to come out of your previous request.");

                await context.PostAsync("However, I can help you with some other popular Support Requests such as: " +
                 "\r\r" + options +                
                        "\r\r Please type any of the above keywords or any other query you may have.");
            }
            catch (Exception)
            {
                await context.PostAsync("Something really bad happened. You can try again later meanwhile I shall check what went wrong.");

                await context.PostAsync("However, I can help you with some other popular Support Requests such as: " +
                 "\r\r" + options +
                        "\r\r Please type any of the above keywords or any other query you may have.");
            }
            finally
            {
               context.Wait(MessageReceived);              

            }
        }        

        #endregion

    }   

}