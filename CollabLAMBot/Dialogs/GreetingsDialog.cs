using BestMatchDialog;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using System;
using System.Threading.Tasks;

namespace CollabLAMBot.Dialogs
{
    [Serializable]
    public class GreetingsDialog : BestMatchDialog<object>
    {
        [BestMatch(new string[] { "Hi", "Hi There", "Hello there", "Hey", "Hello",
        "Hey there", "Greetings", "Good morning", "Good afternoon", "Good evening", "Good day" },
       threshold: 0.5, ignoreCase: true, ignoreNonAlphaNumericCharacters: false)]
        public async Task WelcomeGreeting(IDialogContext context, string messageText)
        {
            //await context.PostAsync("Hello there. How can I help you?");
            //context.Done("true");

            await context.PostAsync("Hi, I am Collab Bot.");
            await Respond(context);

            context.Wait(MessageReceivedAsync);
            //context.Done("true");

        }

        [BestMatch(new string[] { "bye", "bye bye", "got to go", "see you later", "laters", "adios","thanks","thank you" },
            threshold: 0.5, ignoreCase: true, ignoreNonAlphaNumericCharacters: false)]
        public async Task FarewellGreeting(IDialogContext context, string messageText)
        {
            if (messageText.Contains("thank"))
            {
                await context.PostAsync("You are welcome.");
                await context.PostAsync("Have a great day ahead.");
            }
            else
            {
                await context.PostAsync("Thanks for using Collab Bot!");
                await context.PostAsync("Bye. Have a great day ahead.");
            }
            context.Done("true");
        }


        private static async Task Respond(IDialogContext context)
        {
            var userName = String.Empty;
            context.UserData.TryGetValue<string>("Name", out userName);
            if (string.IsNullOrEmpty(userName))
            {
                await context.PostAsync("What is your name?");
                context.UserData.SetValue<bool>("GetName", true);
            }
            else
            {
                await context.PostAsync(String.Format("Hi {0}.  How can I help you today?", userName));
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
            }

            await Respond(context);
            context.Done(message);
        }

        public override async Task NoMatchHandler(IDialogContext context, string messageText)
        {
            context.Done("false");
        }
    }
}