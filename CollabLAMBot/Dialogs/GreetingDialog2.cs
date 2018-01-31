using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using System;
using System.Threading.Tasks;

namespace Avanade.LAM.CollabBOT.Dialogs
{
    [Serializable]
    public class GreetingDialog2 : IDialog
    {
        public async Task StartAsync(IDialogContext context)
        {
            //await Respond(context);

            //var userName = String.Empty;
            //context.UserData.TryGetValue<string>("Name", out userName);
            //if (string.IsNullOrEmpty(userName))
            //{
            //    context.Wait(MessageReceivedAsync);
            //}  

            var options = "* Site Access \n" +
                           "* Site Creation \n" +
                           "* Site Quota Change \n" +
                           "* External User Access \n" +
                           "* Profile Updates \n" +
                           "* User Guides/ KB Articles \n";

            await context.PostAsync("Hi! I am SOHA, your SharePoint Help Assistant \U0001F469 " +
                 "\r\r I can help you with popular Support Requests such as: " +
                 "\r\r" + options +               
                  "\r\r Please type any of the above keywords or any other query you may have.");
            context.Done("");
        }

        private static async Task Respond(IDialogContext context)
        {
            var options = "* Site Access \n"+
                           "* Site Creation \n"+
                           "* Site Quota Change \n"+
                           "* External User Access \n"+
                           "* Profile Updates \n"+
                           "* Knowledge Base Article \n";


            var userName = String.Empty;
            context.UserData.TryGetValue<string>("Name", out userName);
            if (string.IsNullOrEmpty(userName))
            {
                //await context.PostAsync("Hi I'm Collab Bot");
                await context.PostAsync("Hi! Welcome to SharePoint Online Help Desk." +
                "\r\r May I know your name please?");
                context.UserData.SetValue<bool>("GetName", true);
            }
            else
            {
                await context.PostAsync(string.Format("Hi {0} ! ", userName) +
                    " I am here to help you on below options. "+
                        "\r\r" + options+                      
                        "\r\r Please type your question in the space provided below.");

                //await context.PostAsync(String.Format("How can I help you today?"));
                context.Done("Greeting");
            }
        }

        public async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> argument)
        {
            var message = await argument;
            var userName = String.Empty;
            var getName = false;
            context.UserData.TryGetValue<string>("Name", out userName);
            context.UserData.TryGetValue<bool>("GetName", out getName);

            if (getName)
            {
                userName = message.Text;
                context.UserData.SetValue<string>("Name", userName);
                context.UserData.SetValue<bool>("GetName", false);
                await Respond(context);
            }            
            //context.Done("");
        }
    }
}