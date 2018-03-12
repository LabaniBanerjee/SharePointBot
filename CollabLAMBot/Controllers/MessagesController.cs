using Autofac;
using CollabLAMBot.Dialogs;
using CollabLAMBot.Models;
using CollabLAMBot.Utility;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Dialogs.Internals;
using Microsoft.Bot.Connector;
using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;

namespace CollabLAMBot
{
    [BotAuthentication]
    public class MessagesController : ApiController
    {
        /// <summary>
        /// POST: api/Messages
        /// Receive a message from a user and reply to it
        /// </summary>
        public async Task<HttpResponseMessage> Post([FromBody]Activity activity)
        {           

            if (activity.Type == ActivityTypes.Message)
            {
                // call without Luis
                //await Conversation.SendAsync(activity, () => new Dialogs.RootDialog());
                // call with Luis
                await Conversation.SendAsync(activity, () => new CollabBOTLuisDialog());             

            }
            else
            {
                HandleSystemMessage(activity);
            }
            var response = Request.CreateResponse(HttpStatusCode.OK);
            return response;
        }

        

        private Activity HandleSystemMessage(Activity message)
        {
            if (message.Type == ActivityTypes.DeleteUserData)
            {
                // Implement user deletion here
                // If we handle user deletion, return a real message
            }
            else if (message.Type == ActivityTypes.ConversationUpdate)
            {
                // Handle conversation state changes, like members being added and removed
                // Use Activity.MembersAdded and Activity.MembersRemoved and Activity.Action for info
                // Not available in all channels

                // Note: Add introduction here:
                //IConversationUpdateActivity update = message;
                //var client = new ConnectorClient(new Uri(message.ServiceUrl), new MicrosoftAppCredentials());
                //if (update.MembersAdded != null && update.MembersAdded.Any())
                //{
                //    foreach (var newMember in update.MembersAdded)
                //    {
                //        if (newMember.Id != message.Recipient.Id)
                //        {
                //            var reply = message.CreateReply();
                //            //reply.Text = $"Welcome {newMember.Name}!";
                //            string strGreetingContext = UtilityFunctions.setGreetingContext();
                //            string currentLoggedinUser = UtilityFunctions.getWindowsUser();
                //            //reply.Text = $" {strGreetingContext} {currentLoggedinUser} , I am Collab Bot.";
                //            reply.Text = $" {strGreetingContext}, I am Collab Bot.";
                //            client.Conversations.ReplyToActivityAsync(reply);
                //        }
                //    }
                //}

                // Note: Add introduction here:
                //IConversationUpdateActivity update = message;
                //var client = new ConnectorClient(new Uri(message.ServiceUrl), new MicrosoftAppCredentials());
                //if (update.MembersAdded != null && update.MembersAdded.Any())
                //{
                //    foreach (var newMember in update.MembersAdded)
                //    {
                //        if (newMember.Id != message.Recipient.Id)
                //        {
                //            //var reply = message.CreateReply();
                //            ////reply.Text = $"Welcome {newMember.Name}!";
                //            //string strGreetingContext = UtilityFunctions.setGreetingContext();
                //            //string currentLoggedinUser = UtilityFunctions.getWindowsUser();
                //            ////reply.Text = $" {strGreetingContext} {currentLoggedinUser} , I am Collab Bot.";
                //            //reply.Text = $" {strGreetingContext}, I am Collab Bot.";
                //            //client.Conversations.ReplyToActivityAsync(reply);

                //            var options = "* Site Access \n" +
                //           "* Site Creation \n" +
                //           "* Site Quota Change \n" +
                //           "* External User Access \n" +
                //           "* Profile Updates \n" +
                //           "* Knowledge Base Article \n";

                //            var reply = message.CreateReply();
                //            reply.Text = "Hi! My name is SOHA - your SharePoint Online Help Assistant." +
                //                " I am here to help you on below options. " +
                //                 "\r\r" + options +
                //        "\r\r Please type your question in the space provided below.";
                //            client.Conversations.ReplyToActivityAsync(reply);

                //        }
                //    }
                //}

                //using (var scope = Microsoft.Bot.Builder.Dialogs.Internals.DialogModule.BeginLifetimeScope(Conversation.Container, message))
                //{
                //    var client = scope.Resolve<IConnectorClient>();
                //    if (update.MembersAdded.Any())
                //    {
                //        foreach (var newMember in update.MembersAdded)
                //        {
                //            if (newMember.Id != message.Recipient.Id)
                //            {
                //                var reply = message.CreateReply();
                //                var options = "* Site Access \n" +
                //                           "* Site Creation \n" +
                //                           "* Site Quota Change \n" +
                //                           "* External User Access \n" +
                //                           "* Profile Updates \n" +
                //                           "* Knowledge Base Article \n";

                //                //reply.Text = $"Welcome {newMember.Name}!";

                //                reply.Text = "Hi! My name is SOHA - your SharePoint Online Help Assistant." +
                //                                " I am here to help you on below options. " +
                //                                 "\r\r" + options +
                //                        "\r\r Please type your question in the space provided below.";


                //                client.Conversations.ReplyToActivityAsync(reply);
                //            }
                //        }
                //    }
                //}
            }
            else if (message.Type == ActivityTypes.ContactRelationUpdate)
            {
                // Handle add/remove from contact lists
                // Activity.From + Activity.Action represent what happened                
            }
            else if (message.Type == ActivityTypes.Typing)
            {
                // Handle knowing tha the user is typing
            }
            else if (message.Type == ActivityTypes.Ping)
            {
            }

            return null;
        }
    }
}