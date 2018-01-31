using Microsoft.Bot.Builder.Dialogs;
using System;
using System.Threading.Tasks;

namespace Avanade.LAM.CollabBOT.Dialogs
{
    [Serializable]
    public class ThankGreetingDialog : IDialog
    {
        public async Task StartAsync(IDialogContext context)
        {
            await context.PostAsync("You are most welcome! \U0001F642 " + 
                "\r\r Let me know if I can help you with anything else.");
            context.Done("");

        }
    }
}