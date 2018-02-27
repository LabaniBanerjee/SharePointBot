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
            await context.PostAsync("You are most welcome! \U0001F642 ");
            context.Done("");

        }
    }
}