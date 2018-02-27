using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.FormFlow;
using Microsoft.Bot.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace Avanade.LAM.CollabBOT.Dialogs
{
    [Serializable]
    public class ServiceNowDialog : IDialog
    {
        public async Task StartAsync(IDialogContext context)
        { 
            context.Wait(MessageReceivedAsync);
        }

        private async Task MessageReceivedAsync(IDialogContext context, IAwaitable<object> result)
        {
            var message = await result as Activity;

            var options = "* Site Access \n" +
                           "* Site Creation \n" +
                           "* Site Quota Change \n" +
                           "* External User Access \n" +
                           "* Profile Updates \n" +
                           "* User Guides/ KB Articles \n";

            if (message.Text.ToLower().Equals("yes") || message.Text.ToLower().Equals("yep") || message.Text.ToLower().Equals("yup") || message.Text.ToLower().Equals("yeah"))
            {               

                await context.PostAsync("Sure, I can raise a ticket on Service Now, before that I need few inputs from you.");

                var ServiceNowFormDialog = FormDialog.FromForm(this.BuildServiceNowForm, FormOptions.PromptInStart);

                context.Call(ServiceNowFormDialog, this.ResumeServiceNowDialog);
            }
            else
            {
                await context.PostAsync(string.Format("Okay. Let us start over. I can help you with popular Support Requests such as: " +
                       "\r\r" + options +
                       "\r\r Please type any of the above keywords or any other query you may have."));

                context.Done("service now exited with no");

            }
        }

        private async Task ResumeServiceNowDialog(IDialogContext context, IAwaitable<ServiceNowQuery> result)
        {
            try
            {
                var resultFromServiceNow = await result;
                await context.PostAsync($"A ticket has been raised on your behalf. The reference no is RITM"+ DateTime.Now.ToString("yyyyMMddHHmmss"));
                context.Done("service now exited with ticket creation");
            }
            catch (TooManyAttemptsException)
            {
                await context.PostAsync("Sorry \U0001F641 , I am unable to understand you. Let us try again.");
            }
        }

        private IForm<ServiceNowQuery> BuildServiceNowForm()
        {
            OnCompletionAsyncDelegate<ServiceNowQuery> processServiceNow = async (context, authorizationState) =>
            {
                await context.PostAsync($"Thank you! I am creating your request.");
            };

            return new FormBuilder<ServiceNowQuery>()
                .Field(nameof(ServiceNowQuery.ShortDescriptionForTicket))
                .Field(nameof(ServiceNowQuery.ServiceNowPriority))
                .Field(nameof(ServiceNowQuery.ServiceNowCategory))
                .AddRemainingFields()
                .Confirm("Great. I am ready to submit your request with the following details \U0001F447 " +
                        "\r\r Your request states '{ShortDescriptionForTicket}' and you want to get it resolved in '{ServiceNowPriority}' priority " +
                        "\r\r Is that correct?")
                .OnCompletion(processServiceNow)               
                .Build();
        }
    }


    #region query helper

    [Serializable]
    public class ServiceNowQuery
    {
        [Prompt("Please enter some Short Description for your ticket.")]
        public string ShortDescriptionForTicket { get; set; }        
        public Priority? ServiceNowPriority { get; set; }
        public Category? ServiceNowCategory { get; set; }

        public enum Priority
        {
            Low = 1,
            Medium = 2,            
            High = 3         
        }

        public enum Category
        {
            [Describe("Password Support Center")]
            PasswordSupportCenter = 2,
            [Describe("Register New Mobile Device")]
            RegisterNewMobileDevice = 3,
            [Describe("One Drive For Business")]
            OneDriveForBusiness = 4,
            [Describe("Skype For Business")]            
            SkypeForBusiness = 5,
            [Describe("Collaboration Site")]
            CollaborationSite = 6,
            [Describe("Shared Mailbox")]
            SharedMailbox = 7,
            [Describe("Software Installation")]
            SoftwareInstallation =8,
            [Describe("VPN Connectivity")]
            VPNConnectivity =9
        }
    }

    #endregion

}