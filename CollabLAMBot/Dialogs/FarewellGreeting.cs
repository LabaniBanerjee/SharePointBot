using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace Avanade.LAM.CollabBOT.Dialogs
{
    [Serializable]
    public class FarewellGreeting : IDialog
    {
        public async Task StartAsync(IDialogContext context)
        {
           

            //Activity replyToConversation = (Activity)context.MakeMessage();
            //replyToConversation.Attachments = new List<Attachment>();
            //replyToConversation.Attachments.Add(new Attachment()
            //{
            //    //ContentUrl = /*"http://i.giphy.com/p3BDz27c5RlIs.gif",*/ "https://media1.tenor.com/images/b9371273ae94a946e92074d1b9696680/tenor.gif?itemid=10897308", https://media.giphy.com/media/ZIMHexjRxlP0I/giphy.gif
            //    ContentUrl = "https://media.giphy.com/media/ZIMHexjRxlP0I/giphy.gif",
            //    ContentType = "image/gif",
                
            //});

            //await context.PostAsync(replyToConversation);

            await context.PostAsync("Thanks for using SharePoint Help Assistant!" +
               "\r\r Bye \U0001F44B . Have a great day ahead.");
            context.Done("Farewell");

        }
    }
}