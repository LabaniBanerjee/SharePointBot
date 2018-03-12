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

                try
                {
                    SharePointPrimary obj = new SharePointPrimary();
                    Dictionary<ListItem, string> searchedDocs = new Dictionary<ListItem, string>();

                    if (_strhelpArticle.ToLower().Contains("video"))
                    {
                        await context.PostAsync($"I have found one video on your search topic \U0001F44D ");

                        var videoCard = new VideoCard
                        {
                            Title = "Build a Chat Bot with Azure Bot Service",
                            Subtitle = "by ShaunLuttin, Anthony Chu",
                            Text = "Microsoft MVPs Anthony Chu and Shaun Luttin sit down and build a natural language chat bot from scratch using the new Azure Bot Service and Microsoft Cognitive Services' LUIS (Language Understanding Intelligent Service).",
                            Image = new ThumbnailUrl
                            {
                                Url = "http://blog.legalsolutions.thomsonreuters.com/wp-content/uploads//2015/01/online-bot.png",
                                Alt = "Azure Bot"
                            },
                            Media = new List<MediaUrl>
                            {
                                new MediaUrl()
                                {
                                    Url = "https://sec.ch9.ms/ch9/a7a4/10df13cf-a7ac-40a2-b713-6fcc935ba7a4/buildachatbotwithazurebotservice_high.mp4"
                                }
                            },
                            Buttons = new List<CardAction>
                            {
                                new CardAction()
                                {
                                    Title = "Learn More",
                                    Type = ActionTypes.OpenUrl,
                                    Value = "https://channel9.msdn.com/Blogs/MVP-Azure/build-a-chatbot-with-azure-bot-service"
                                }
                            }
                        };

                        Microsoft.Bot.Connector.Attachment attachment = new Microsoft.Bot.Connector.Attachment();
                        var replyMessage = context.MakeMessage();
                        attachment = videoCard.ToAttachment();
                        replyMessage.Attachments.Add(attachment);

                        await context.PostAsync(replyMessage);
                        context.Done("Done");
                    }
                    else
                    {
                        searchedDocs = obj.SearchHelpArticleonTopic(_strhelpArticle);
                        if (searchedDocs.Count > 0)
                        {
                            if (searchedDocs.Count == 1)
                                await context.PostAsync($"I have found {searchedDocs.Count} article on your search topic \U0001F44D ");
                            else
                                await context.PostAsync($"I have found {searchedDocs.Count} articles on your search topic \U0001F44D ");

                            foreach (var eachDoc in searchedDocs)
                            //foreach (ListItem eachDoc in searchedDocs)
                            {
                                Microsoft.Bot.Connector.Attachment attachment = new Microsoft.Bot.Connector.Attachment();
                                //attachment.ContentType = "application/pdf";
                                //attachment.ContentUrl = eachDoc.Value;
                                //attachment.Name = eachDoc.Key;

                                var replyMessage = context.MakeMessage();
                                //replyMessage.Attachments = new List<Attachment> { attachment };
                                //await context.PostAsync(replyMessage);
                                                               

                                var heroCard = new HeroCard
                                {
                                    Title = eachDoc.Key.DisplayName,
                                    Subtitle = "by Avanade Collab SL Capability",
                                    Text = eachDoc.Key.FieldValuesAsHtml["KpiDescription"],
                                    //Images = new List<CardImage> { new CardImage("https://sec.ch9.ms/ch9/7ff5/e07cfef0-aa3b-40bb-9baa-7c9ef8ff7ff5/buildreactionbotframework_960.jpg") },
                                    //Images = new List<CardImage> {new CardImage (eachDoc.Key.FieldValuesAsHtml["ThumbnailURL"]) },
                                    Images = new List<CardImage> { new CardImage(eachDoc.Value.ToString()) },
                                    Buttons = new List<CardAction> { new CardAction(ActionTypes.OpenUrl, "Read more", value: Constants.RootSiteCollectionURL + eachDoc.Key.File.ServerRelativeUrl) }
                                };


                                attachment = heroCard.ToAttachment();
                                replyMessage.Attachments.Add(attachment);

                                await context.PostAsync(replyMessage);

                            }
                           
                            context.Done("Done");
                        }
                        else
                        {
                            await context.PostAsync($"No document found. Do you want to search for any other topic?");
                            //context.Done("Not Done");
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
                context.Done("Not Done");
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
    }
}