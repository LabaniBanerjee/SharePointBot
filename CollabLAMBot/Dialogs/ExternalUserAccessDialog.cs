using Microsoft.Bot.Builder.Dialogs;
using System;
using System.Threading.Tasks;


namespace CollabLAMBot.Dialogs
{
    [Serializable]
    public class ExternalUserAccessDialog : IDialog
    {
        public async Task StartAsync(IDialogContext context)
        {
            await context.PostAsync("Sorry, I am unable to grant access to any External User at the moment. Please select another option. ");
            context.Done("false");
            //context.Wait(MessageReceived);
        }

        private async Task MessageReceived(IDialogContext context, IAwaitable<object> result)
        {
            var myresult = await result;
            context.Done("ExternalUserAccessDialog says : " + myresult);
        }
    }
}