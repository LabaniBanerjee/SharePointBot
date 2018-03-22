using Avanade.LAM.CollabBOT.Utility;
using HtmlAgilityPack;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.FormFlow.Advanced;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System;
using System.Threading.Tasks;
using Microsoft.Bot.Builder.FormFlow;
using CollabLAMBot.LAM;
using Microsoft.Bot.Connector;
using Microsoft.SharePoint.Client;
using Avanade.LAM.CollabBOT.LAM;
using System.Net;

namespace Avanade.LAM.CollabBOT.Dialogs
{
    [Serializable]
    public class ArticleSearch : IDialog<string>
    {
        #region commented code
        //const string BLOG_URI = "http://ankitbko.github.io/archive";

        //public List<Post> GetPostsWithTag(string tag)
        //{
        //    List<Post> posts = GetAllPosts();
        //    return posts.FindAll(p => IsMatch(tag, p.Tags));
        //}

        //public List<Post> GetAllPosts()
        //{
        //    var blogHtml = GetHtmlFromBlog();
        //    return blogHtml.DocumentNode.SelectSingleNode("//ul").ChildNodes.Where(t => t.Name == "li")     // Select all posts
        //        .Select((f, n) =>
        //            new Post()
        //            {
        //                Name = f.SelectSingleNode("a").InnerText,                                           // Select post name
        //                Tags = f.SelectSingleNode("ul").ChildNodes.Where(t => t.Name == "li")               // Select Tags
        //                        .Select(t => t.InnerText).ToList(),                                         // Select tag name
        //                Url = f.SelectSingleNode("a[@href]").GetAttributeValue("href", string.Empty) //Select URL
        //            }).ToList();

        //}

        //private HtmlDocument GetHtmlFromBlog()
        //{
        //    HtmlWeb web = new HtmlWeb();
        //    return web.Load(BLOG_URI);
        //}

        //private bool IsMatch(string tag, List<string> tags)
        //{
        //    var terms = tags.SelectMany(t => Language.GenerateTerms(Language.CamelCase(t), 3));
        //    foreach (var t in terms)
        //    {
        //        if (Regex.IsMatch(tag, t))
        //            return true;
        //    }
        //    return false;
        //}
        #endregion

        public string _intent;

        public ArticleSearch()
        {

        }
        public ArticleSearch(string intent)
        {
            this._intent = intent;
        }
        public async Task StartAsync(IDialogContext context)
        {
            //await context.PostAsync("Welcome to Article Search Dialog!");

            var articleSearchFormDialog = FormDialog.FromForm(this.BuildSearchArticleForm, FormOptions.PromptInStart);

            context.Call(articleSearchFormDialog, this.ResumeArticleSearchDialog);
        }       

        private async Task ResumeArticleSearchDialog(IDialogContext context, IAwaitable<ArticleSearchQuery> result)
        {
            try
            {
                var resultFromArticleSearch = await result;                
                string _strhelpArticle = Convert.ToString(resultFromArticleSearch.HelpArticle);
                int _iLearningMaterial = Convert.ToInt32(resultFromArticleSearch.LearningMaterial);

                try
                {
                    SharePointPrimary obj = new SharePointPrimary();
                    Dictionary<ListItem, string> searchedDocs = new Dictionary<ListItem, string>();

                    if (_iLearningMaterial == 2)
                    {  
                        searchedDocs = obj.SearchLearningVideoByTopic(_strhelpArticle);
                        if (searchedDocs.Count > 0)
                        {
                            if(searchedDocs.Count == 1)
                                await context.PostAsync($"I have found {searchedDocs.Count} video on your search topic \U0001F44D ");
                            else
                                await context.PostAsync($"I have found {searchedDocs.Count} videos on your search topic \U0001F44D ");

                            foreach (var eachDoc in searchedDocs)
                            {
                                try
                                {
                                    var videoCard = new VideoCard();
                                    videoCard.Title = eachDoc.Key.FieldValuesAsHtml["Title"];
                                    videoCard.Subtitle = eachDoc.Key.FieldValuesAsHtml["SubTitle"];
                                    videoCard.Text = eachDoc.Key.FieldValuesAsHtml["VideoSetDescription"];

                                    string _previewImageBase64 = obj.GetImage(eachDoc.Key.FieldValuesAsText["AlternateThumbnailUrl"].Split(',')[0].ToString(), _iLearningMaterial);


                                    if (!string.IsNullOrEmpty(_previewImageBase64))
                                    {
                                        videoCard.Image = new ThumbnailUrl(
                                            url: _previewImageBase64,
                                            alt: "Learning Video");
                                    }

                                    //string _videoBase64 = obj.GetVideo("https://avaindcollabsl.sharepoint.com/sites/SOHA_HelpRepository/VideoRepository/Build%20a%20Chat%20Bot%20with%20Azure%20Bot%20Service/buildachatbotwithazurebotservice_high.mp4");

                                    //string _videoBase64 = obj.GetVideo(eachDoc.Value);


                                    if (!string.IsNullOrEmpty(eachDoc.Value))
                                    {
                                        videoCard.Media = new List<MediaUrl> { new MediaUrl("https://sec.ch9.ms/ch9/a7a4/10df13cf-a7ac-40a2-b713-6fcc935ba7a4/buildachatbotwithazurebotservice.mp4") };

                                        videoCard.Buttons = new List<CardAction> { new CardAction(
                                        type  : ActionTypes.OpenUrl,
                                        title : "Learn more",
                                        value : "https://channel9.msdn.com/Blogs/MVP-Azure/build-a-chatbot-with-azure-bot-service") };
                                    }


                                    //if (!string.IsNullOrEmpty(/*eachDoc.Value*/_videoBase64))
                                    //{
                                    //    videoCard.Media = new List<MediaUrl> { new MediaUrl(_videoBase64) };

                                    //    videoCard.Buttons = new List<CardAction> { new CardAction(
                                    //    type  : ActionTypes.OpenUrl,
                                    //    title : "Learn more",
                                    //    value : eachDoc.Value) };
                                    //}

                                    Microsoft.Bot.Connector.Attachment attachment = new Microsoft.Bot.Connector.Attachment();
                                    var replyMessage = context.MakeMessage();
                                    attachment = videoCard.ToAttachment();
                                    replyMessage.Attachments.Add(attachment);

                                    await context.PostAsync(replyMessage);
                                }
                                catch(Exception ex)
                                {
                                    await context.PostAsync(ex.Message);
                                }
                            }

                            await context.PostAsync("Hope you have found the result informative. Do you want to search for any other topic?");
                            context.Wait(MessageReceived);
                        }
                        else
                        {
                            await context.PostAsync($"No video found. Do you want to search for any other topic?");
                            context.Wait(MessageReceived);
                        }                       

                    }
                    else if(_iLearningMaterial == 1)
                    {
                        searchedDocs = obj.SearchHelpArticleonTopic(_strhelpArticle);
                        if (searchedDocs.Count > 0)
                        {
                            if (searchedDocs.Count == 1)
                                await context.PostAsync($"I have found {searchedDocs.Count} article on your search topic \U0001F44D ");
                            else
                                await context.PostAsync($"I have found {searchedDocs.Count} articles on your search topic \U0001F44D ");

                            foreach (var eachDoc in searchedDocs)                          
                            {
                                Microsoft.Bot.Connector.Attachment attachment = new Microsoft.Bot.Connector.Attachment();                                

                                var replyMessage = context.MakeMessage();                                                           

                                var heroCard= new HeroCard();

                                string _previewImageBase64 = obj.GetImage(eachDoc.Value.ToString(), _iLearningMaterial);

                                heroCard.Title = eachDoc.Key.DisplayName;
                                heroCard.Subtitle = "by Avanade Collab SL Capability";
                                heroCard.Text = eachDoc.Key.FieldValuesAsHtml["KpiDescription"];
                                heroCard.Buttons = new List<CardAction> { new CardAction(ActionTypes.OpenUrl, "Read more", value: Constants.RootSiteCollectionURL + eachDoc.Key.File.ServerRelativeUrl) };

                                if(!string.IsNullOrEmpty(_previewImageBase64))
                                    heroCard.Images = new List<CardImage> { new CardImage(_previewImageBase64) };                               

                                
                                attachment = heroCard.ToAttachment();
                                replyMessage.Attachments.Add(attachment);

                                await context.PostAsync(replyMessage);
                            }

                            await context.PostAsync("Hope you have found the result informative. Do you want to search for any other topic?");
                            context.Wait(MessageReceived);
                            
                        }
                        else
                        {
                            await context.PostAsync($"No document found. Do you want to search for any other topic?");                           
                            context.Wait(MessageReceived);                            
                        }
                    }                   
                }
                catch(TooManyAttemptsException)
                {
                    context.Fail(new TooManyAttemptsException("Unable to find any help document."));
                }               
            }
            catch (TooManyAttemptsException)
            {
                await context.PostAsync("Sorry \U0001F641 , I am unable to understand you. Let us try again.");
            }
        }


        private async Task MessageReceived(IDialogContext context, IAwaitable<object> result)
        {
            var message = await result as Activity;
            

            if (message.Text.ToLower().Equals("yes") || message.Text.ToLower().Equals("yep") 
                || message.Text.ToLower().Equals("yup") || message.Text.ToLower().Equals("yeah")
                || message.Text.ToLower().Equals("y"))
            {
                var articleSearchFormDialog = FormDialog.FromForm(this.BuildSearchArticleForm, FormOptions.PromptInStart);

                context.Call(articleSearchFormDialog, this.ResumeArticleSearchDialog);               
            }
            else
            {
                context.Done("Article Done");
            }

        }

        private IForm<ArticleSearchQuery> BuildSearchArticleForm()
        {
            OnCompletionAsyncDelegate<ArticleSearchQuery> processArticleSearch = async (context, authorizationState) =>
            {
                await context.PostAsync($"Searching knowledge base for you...");
            };

            return new FormBuilder<ArticleSearchQuery>()
                .Field(nameof(ArticleSearchQuery.HelpArticle))
                .AddRemainingFields()
                .Confirm("Sure. So you need help on '{HelpArticle}'" +
                        "\r\r Is that correct?")
                .OnCompletion(processArticleSearch)
                //.Message("")
                .Build();
        }
    }

    [Serializable]
    public class ArticleSearchQuery
    {
        [Prompt("Which topic do you want me to search for?")]
        public string HelpArticle { get; set; }

        public UserGuideType? LearningMaterial { get; set; }

        public enum UserGuideType
        {
            Article = 1,
            Video = 2
        }
    }
}