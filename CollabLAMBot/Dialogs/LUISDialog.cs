using CollabLAMBot.Models;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.FormFlow;
using Microsoft.Bot.Builder.Luis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace CollabLAMBot.Dialogs
{
    [LuisModel("0813b523-b87f-451c-b8d2-3c6224154886", "c428e39842534c96b7cb10688ba2aa2e")]
    [Serializable]
    public class LUISDialog : LuisDialog<object>
    {
        //private readonly BuildFormDelegate<SiteAccessDialog> AuthorizeUser;

        //public LUISDialog(BuildFormDelegate<SiteAccessDialog> authorizeuser)
        //{
        //    this.AuthorizeUser = authorizeuser;
        //}

        public LUISDialog(params ILuisService[] services) : base(services)
        {
        }

        [LuisIntent("None")]
        [LuisIntent("")]
        public async Task None(IDialogContext context , LuisRequest result)
        {
            var options = "* Site Access \n" +
                           "* Site Creation \n" +
                           "* Site Quota Change \n" +
                           "* External User Access \n" +
                           "* Profile Updates \n" +
                           "* Knowledge Base Article \n";

            await context.PostAsync("Sorry, I am unable to understand you. Let us start over."+
                 "\r\r" + options +
                        "\r\r Please type your question in the space provided below.");
            context.Wait(MessageReceived);
        }


        [LuisIntent("UserAuthorization")]
        public async Task UserAuthorization(IDialogContext context, LuisRequest result)
        {
            //context.Call(new SiteAccessDialog(),SiteAccessCallback);
            //var siteAccess = new FormDialog<SiteAccessDialog>(new SiteAccessDialog(),
            //    this.AuthorizeUser, FormOptions.PromptInStart);

            //context.Call<SiteAccessDialog>(siteAccess, SiteAccessCallback);

            try
            {
                await context.PostAsync("Weclome to SPOAssistant Site Access");
                var spoForm = new FormDialog<SPOAssistant>(new SPOAssistant(), SPOAssistant.BuildForm, FormOptions.PromptInStart);
                context.Call(spoForm, SPOAssistantFormComplete);
            }
            catch (Exception)
            {
                await context.PostAsync("Something really bad happened. You can try again later meanwhile I'll check what went wrong.");
                context.Wait(MessageReceived);
            }


        }

        private async Task SPOAssistantFormComplete(IDialogContext context, IAwaitable<SPOAssistant> result)
        {        

            try
            {
                //var feedback = await result;
                //string message = GenerateEmailMessage(feedback);
                //var success = await EmailSender.SendEmail(recipientEmail, senderEmail, $"Email from {feedback.Name}", message);
                //if (!success)
                //    await context.PostAsync("I was not able to send your message. Something went wrong.");
                //else
                //{
                //    await context.PostAsync("Thanks for the feedback.");
                //    await context.PostAsync("What else would you like to do?");
                //}

                var resultFromSiteAccess = await result;

                await context.PostAsync($"SPOAssistant dialog just told me this: {resultFromSiteAccess}");  

            }
            catch (FormCanceledException)
            {
                await context.PostAsync("Don't want to send feedback? That's ok. You can drop a comment below.");
            }
            catch (Exception)
            {
                await context.PostAsync("Something really bad happened. You can try again later meanwhile I'll check what went wrong.");
            }
            finally
            {
                context.Wait(MessageReceived);
            }



        }

    }
}