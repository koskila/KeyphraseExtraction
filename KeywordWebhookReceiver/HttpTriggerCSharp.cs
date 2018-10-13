using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using OfficeDevPnP.Core;
using Microsoft.SharePoint.Client;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Microsoft.Azure.CognitiveServices.Language.TextAnalytics.Models;
using Microsoft.Azure.CognitiveServices.Language.TextAnalytics;
using System.Globalization;
using Newtonsoft.Json;
using Microsoft.SharePoint.Client.Taxonomy;

using TaxonomyExtensions = Microsoft.SharePoint.Client.TaxonomyExtensions;

namespace KeywordWebhookReceiver
{
    public static class HttpTriggerCSharp
    {
        public static readonly string siteUrl = System.Configuration.ConfigurationManager.AppSettings["SiteUrl"];
        // These keys are for posti@koskila.net
        public static readonly string key1 = System.Configuration.ConfigurationManager.AppSettings["CognitiveServicesAPIkey"];

        public static readonly string userName = System.Configuration.ConfigurationManager.AppSettings["SPO_UserName"];
        public static readonly string password = System.Configuration.ConfigurationManager.AppSettings["SPO_Password"];
        public static readonly string listName = System.Configuration.ConfigurationManager.AppSettings["SPO_ListName"];

        public static readonly int _accuracyLevel = int.Parse(System.Configuration.ConfigurationManager.AppSettings["AccuracyLevel"]);

        /// <summary>
        /// If you uncomment the row below, the cognitive services part of the code will revert to test/dev mode
        /// </summary>
        private static readonly string[] terms = new string[] { };

        private static int lcid = 1033;

        /// <summary>
        /// Guid of the termset used by our custom Keywords -field
        /// </summary>
        private static Guid wantedGuid = new Guid("b194954e-ba65-4a51-a5b8-c4f732573d24"); 
        private static int keywordCount = 10;

        [FunctionName("HttpTriggerCSharp")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            string message = "";
            string str = "C# HTTP trigger function processed a request.";
            log.Info("====================================");
            log.Info(str);
            message += str;

            List<string> keyPhrases = new List<string>();

            string description = "";

            // Get request body
            dynamic data = await req.Content.ReadAsAsync<object>();

            string id = data.ID;

            log.Info("Document id: " + id);
            message += "Document id: " + id;

            OfficeDevPnP.Core.AuthenticationManager authManager = new OfficeDevPnP.Core.AuthenticationManager();
            try
            {
                // Connects to SharePoint online site
                using (var ctx = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, userName, password))
                {
                    // List Name input  
                    // Retrieves list object using title  
                    List list = ctx.Site.RootWeb.GetListByTitle(listName);
                    if (list != null)
                    {
                        // Returns required result  
                        ListItem li = list.GetItemById(id);

                        ctx.Load(li);
                        ctx.Load(li.File);
                        ctx.ExecuteQuery();

                        ctx.ExecuteQuery();

                        // We CAN extract text out of most documents with the library, but for this demo I'm limiting our options to these 2 that I know to be working :)
                        if (li.File.Name.IndexOf(".pdf") >= 0 || li.File.Name.IndexOf(".doc") >= 0)
                        {
                            li.File.OpenBinaryStream();

                            ctx.Load(li.File);
                            ctx.ExecuteQuery();

                            log.Info("It was a valid file! Continuing into handling...");

                            
                            try
                            {
                                log.Info("Got a file! Name: " + li.File.Name);


                                var fileRef = li.File.ServerRelativeUrl;
                                var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, fileRef);
                                fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, fileRef);

                                using (var ms = new MemoryStream())
                                {
                                    log.Info("Extracting text..");

                                    fileInfo.Stream.CopyTo(ms);
                                    byte[] fileContents = ms.ToArray();

                                    var extractor = new TikaOnDotNet.TextExtraction.TextExtractor();
                                    var extractionResult = extractor.Extract(fileContents);
                                    string text = extractionResult.Text;

                                    List<MultiLanguageInput> analyzable = FormatAnalyzableText(ref text);

                                    log.Info("Formed altogether " + analyzable.Count + " sentences to analyze!");

                                    int snippetEnd = 500 < text.Length ? 500 : text.Length;
                                    log.Info("Extracted text! First few rows here.. \r\n " + text.Substring(0, snippetEnd));

                                    
                                    RunTextAnalysis(ref keyPhrases, analyzable, log);
                                    
                                }

                                log.Info("Found " + keyPhrases.Count + " key phrases! First 20 are here: ");
                                foreach (var kp in keyPhrases.Take(20))
                                {
                                    log.Info(kp);
                                }
                                    

                                try
                                {
                                    log.Info("Saving to SharePoint..");
                                    TextInfo ti = new CultureInfo("en-US", false).TextInfo;

                                    li["Title"] = ti.ToTitleCase(li.File.Name);

                                    // then write the most important keyphrases back
                                    foreach (var s in keyPhrases.Take(keywordCount))
                                    {
                                        description += s + "\r\n";
                                    }

                                    li.Update();

                                    try
                                    {
                                        ctx.Load(list.Fields);
                                        ctx.ExecuteQuery();

                                        log.Info("Updating Managed Metadata...");

                                        var fieldnames = new string[] { "Keywords" };
                                        var field = list.GetFields(fieldnames).First();

                                        // setting managed metadata
                                        log.Info("Updating keywords to taxonomy! Taking: " + keywordCount);
                                        UpdateManagedMetadata(keyPhrases.Take(keywordCount).ToArray(), log, ctx, li, field, wantedGuid);

                                        ctx.ExecuteQuery();
                                    }
                                    catch (Exception ex)
                                    {
                                        log.Error(ex.Message);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    log.Error(ex.Message);
                                }
                            }
                            catch (Exception ex)
                            {
                                log.Error(ex.Message);
                                return req.CreateResponse(HttpStatusCode.InternalServerError, ex.Message);
                            }
                            
                        }
                        else
                        {
                            return req.CreateResponse(HttpStatusCode.OK, "File was not pdf or doc");
                        }

                        return req.CreateResponse(HttpStatusCode.OK, list.Id);
                    }
                    else
                    {
                        log.Info("List is not available on the site");
                        return req.CreateResponse(HttpStatusCode.NotFound, "List is not available on the site");
                    }
                }
            }
            catch (Exception ex)
            {
                log.Info("Error Message: " + ex.Message);
                message += "Error Message: " + ex.Message;
            }

            log.Info("");

            var returnable = JsonConvert.SerializeObject(keyPhrases);

            return keyPhrases.Count <= 0
                ? req.CreateResponse(HttpStatusCode.BadRequest, "Couldn't analyze file. Please verify the POST payload!")
                : req.CreateResponse(HttpStatusCode.OK, returnable);
        }

        private static List<MultiLanguageInput> FormatAnalyzableText(ref string text)
        {
            // sanitize text a bit
            text = Regex.Replace(text, @"[\r\n\t\f\v]", " ");
            // remove extremely long words - they'll be headers, malformed parts or urls
            text = Regex.Replace(text, @"\S{30,}", " ", RegexOptions.None);
            // remove numbers and everything else but text.
            text = Regex.Replace(text, @"[^a-zA-Z.,'!?הצוü]", " ", RegexOptions.IgnoreCase);
            // lastly, remove extra whitespace
            text = Regex.Replace(text, @"( +)", " ");

            List<string> sentences = new List<string>();
            var RegEx_SentenceDelimiter = new Regex(@"(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?)\s");
            sentences = RegEx_SentenceDelimiter.Split(text).ToList();

            // figure out, which sentence length we're using based on set accuracylevel. The default value is 5120 (set by the API)
            int limit;
            if (_accuracyLevel == 0) limit = 5120;
            else if (_accuracyLevel == 1) limit = 2560;
            else if (_accuracyLevel == 2) limit = 640; // please note, that this causes a roughly 8-fold increase in quota consumption!
            else
            {
                throw new ArgumentOutOfRangeException("AccuracyLevel", "Internal parameter _accuracyLevel was not valid. Expected (int)0-2, but was " + _accuracyLevel + ". Please check your Application Settings (properties) configuration, or settings.json file!");
            }

            List<string> finalizedSentences = new List<string>();

            string sentenceCandidate = "";
            foreach (var sentence in sentences)
            {
                // SANITIZE AND SPLIT

                // drop short sentences (they'll be like "et al", one-liners like "go figure" or just "."
                if (sentence.Length < 10) continue;

                // combine or add other sentences
                if (sentenceCandidate.Length + sentence.Length > limit)
                {
                    finalizedSentences.Add(sentenceCandidate);
                    sentenceCandidate = sentence;
                }
                else
                {
                    sentenceCandidate += " " + sentence;
                }
            }
            // finally, add the last candidate
            finalizedSentences.Add(sentenceCandidate);

            var analyzable = new List<MultiLanguageInput>();

            int i = 0;
            foreach (var s in finalizedSentences)
            {
                if (s.Length > 10) analyzable.Add(new MultiLanguageInput("en", i + "", s));
                i++;
            }

            return analyzable;
        }

        /// <summary>
        /// Adds found key phrases to references list of keyPhrases - the most important ones first.
        /// </summary>
        /// <param name="keyPhrases"></param>
        /// <param name="analyzable"></param>
        /// <param name="log"></param>
        private static void RunTextAnalysis(ref List<string> keyPhrases, List<MultiLanguageInput> analyzable, TraceWriter log)
        {
            var batch = new MultiLanguageBatchInput();
            batch.Documents = analyzable;

            ITextAnalyticsAPI client = new TextAnalyticsAPI();
            client.AzureRegion = AzureRegions.Westus;
            client.SubscriptionKey = key1;

            try
            {
                var result = client.KeyPhrases(batch);

                foreach (var row in result.Documents)
                {
                    foreach (var kp in row.KeyPhrases)
                    {
                        keyPhrases.Add(kp);
                    }
                }
            }
            catch (Exception ex)
            {
                log.Warning(ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="values"></param>
        /// <param name="log"></param>
        /// <param name="ctx"></param>
        /// <param name="item"></param>
        /// <param name="field"></param>
        private static void UpdateManagedMetadata(string[] values, TraceWriter log, ClientContext ctx, ListItem item, Field field, Guid guid)
        {
            try
            {
                var taxSession = ctx.Site.GetTaxonomySession();
                var terms = new List<KeyValuePair<Guid, string>>();

                var store = TaxonomyExtensions.GetDefaultKeywordsTermStore(ctx.Site);
                //var keywordTermSet = store.KeywordsTermSet;
                var keywordTermSet = store.GetTermSet(guid);

                ctx.Load(field);
                ctx.ExecuteQuery();

                TaxonomyField taxField = ctx.CastTo<TaxonomyField>(field);
                taxField.EnsureProperty(tf => tf.AllowMultipleValues);

                ctx.Load(taxSession);
                ctx.Load(store);
                ctx.Load(taxField);
                ctx.Load(keywordTermSet);
                ctx.Load(item);
                ctx.ExecuteQuery();

                taxField.IsKeyword = false;
                //taxField.TermSetId = keywordTermSet.Id;
                taxField.Update();

                ctx.Load(taxField);
                ctx.ExecuteQuery();

                foreach (var arrayItem in values)
                {
                    Term term1 = null;
                    Guid termGuid = Guid.Empty;
                    term1 = keywordTermSet.Terms.GetByName(arrayItem);

                    // Test if this term is available
                    try
                    {
                        ctx.Load(term1);
                        ctx.ExecuteQuery();
                    }
                    catch (Exception ex)
                    {
                        log.Info("Encountered an error, probably because a term didn't exist yet! Creating...");

                        term1 = keywordTermSet.CreateTerm(arrayItem, lcid, Guid.NewGuid());
                        term1.IsAvailableForTagging = true;
                    }                        
                    
                    ctx.Load(term1);
                    ctx.ExecuteQuery();

                    store.CommitAll();
                    ctx.ExecuteQuery();

                    terms.Add(new KeyValuePair<Guid, string>(term1.Id, term1.Name));
                }

                ctx.Load(item);
                ctx.ExecuteQuery();
                
                TaxonomyExtensions.SetTaxonomyFieldValues(item, taxField.Id, terms);
               
                item.Update();
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
            }
        }
    }
}
