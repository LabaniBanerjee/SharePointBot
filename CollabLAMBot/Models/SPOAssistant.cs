using Microsoft.Bot.Builder.FormFlow;
using System;

namespace CollabLAMBot.Models
{
    public enum APIType
    {
        PowerShell,
        CSOM
    }

    public enum TicketTypes
    {        
        UserAuthorization,
        SiteCreation,
        SiteQuotaChange,
        ExternalUserAccess,
        ProfileUpdates
    }

    [Serializable]
    public class SPOAssistant
    {
        
        public TicketTypes? SPOTicketType;
        public string siteCollectionURL;

        [Prompt(new string[] { "What is the site collection url?" })]
        public string SiteCollectionURL { get; set; }

        [Prompt("What is your spo user id?")]
        public string SPOUserID { get; set; }

        


        public static IForm<SPOAssistant>BuildForm()
        {
            return new FormBuilder<SPOAssistant>()
                .Message("Welcome to SharePoint Online Assistant Bot ! .")
                .Field(nameof(SiteCollectionURL))
                .Field(nameof(SPOUserID))                
                .Build();

        }
    }
}