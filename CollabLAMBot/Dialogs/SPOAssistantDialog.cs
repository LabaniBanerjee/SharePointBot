using CollabLAMBot.Models;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.FormFlow;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;

namespace CollabLAMBot.Dialogs
{
    public class SPODialog 
    {
        public static readonly IDialog<string> dialog = Chain.PostToChain()
            .Select(msg => msg.Text)
            .Switch(
                     new RegexCase<IDialog<string>>(new Regex("^hi", RegexOptions.IgnoreCase), (context, txt) =>
                     {
                         return Chain.ContinueWith(new RootDialog(), AfterGreetingContinuation);
                     }),
                    new RegexCase<IDialog<string>>(new Regex("^hello", RegexOptions.IgnoreCase), (context, txt) =>
                    {
                        return Chain.ContinueWith(new RootDialog(), AfterGreetingContinuation);
                    }),
                    new RegexCase<IDialog<string>>(new Regex("^hey", RegexOptions.IgnoreCase), (context, txt) =>
                    {
                        return Chain.ContinueWith(new RootDialog(), AfterGreetingContinuation);
                    }),
                     new DefaultCase<string, IDialog<string>>((context, txt) =>
                     {
                         return Chain.ContinueWith(FormDialog.FromForm(SPOAssistant.BuildForm, FormOptions.PromptFieldsWithValues), AfterGreetingContinuation);
                     }))
         .Unwrap()
         .PostToUser();

        private async static Task<IDialog<string>> AfterGreetingContinuation(IBotContext context, IAwaitable<object> res)
        {
            var token = await res;
            var name = "User";
            context.UserData.TryGetValue<string>("Name", out name);
            return Chain.Return($"Thank you for using SharePoint Online Assistent bot : {name}");
        }
    }
}