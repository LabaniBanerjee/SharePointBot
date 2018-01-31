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
                    Dictionary<string, string> searchedDocs = new Dictionary<string, string>();

                    searchedDocs = obj.SearchHelpArticleonTopic(1);
                    if(searchedDocs.Count>0)
                    {
                        await context.PostAsync($"I have found {searchedDocs.Count} article(s) on your search topic \U0001F44D ");
                       
                        foreach (var eachDoc in searchedDocs)
                        {                            
                            Attachment attachment = new Attachment();
                            attachment.ContentType = "application/pdf";
                            attachment.ContentUrl = eachDoc.Value;
                            attachment.Name = eachDoc.Key;

                            var replyMessage = context.MakeMessage();
                            replyMessage.Attachments = new List<Attachment> { attachment };
                            await context.PostAsync(replyMessage);
                        }
                        context.Done("article found");
                    }
                    else
                    {
                        await context.PostAsync($"No document found \U0001F641 ");
                        context.Done("article not found");
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