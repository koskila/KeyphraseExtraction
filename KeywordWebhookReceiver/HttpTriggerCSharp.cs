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
using Microsoft.SharePoint.Client.Taxonomy;

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


        [FunctionName("HttpTriggerCSharp")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            string message = "";
            string str = "C# HTTP trigger function processed a request.";
            log.Info("====================================");
            log.Info(str);
            message += str;

            // parse query parameter
            string name = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "name", true) == 0)
                .Value;

            string id = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "id", true) == 0)
                .Value;

            string link = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "link", true) == 0)
                .Value;

            List<string> keyPhrases = new List<string>();

            // Get request body
            dynamic data = await req.Content.ReadAsAsync<object>();

            try
            {
                link = data.link;
            }
            catch (Exception ex)
            { 
                log.Info(ex.Message);
                //throw;
            }

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

                        Microsoft.SharePoint.Client.File file = li.File;
                        ctx.Load(file);
                        ctx.ExecuteQuery();

                        var terms = new string[] { "test term", "lol test" };
                        keyPhrases.AddRange(terms);

                        if (li.File.Name.IndexOf("pdf") >= 0)
                        {
                            li.File.OpenBinaryStream();

                            ctx.Load(li.File);
                            ctx.Load(file);
                            ctx.ExecuteQuery();

                            log.Info("It was a pdf! Continuing into handling...");

                            using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                            {
                                try
                                {
                                    log.Info(file.Name);

                                    var fileRef = li.File.ServerRelativeUrl;
                                    var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, fileRef);
                                    fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, fileRef);

                                    using (var ms = new MemoryStream())
                                    {
                                        fileInfo.Stream.CopyTo(ms);
                                        byte[] fileContents = ms.ToArray();

                                        var extractor = new TikaOnDotNet.TextExtraction.TextExtractor();
                                        var extractionResult = extractor.Extract(fileContents);
                                        string text = extractionResult.Text;

                                        // sanitize text a bit
                                        text = Regex.Replace(text, @"[\r\n\t\f\v]", " ");
                                        text = Regex.Replace(text, @"[^a-z.,!?]", " ", RegexOptions.IgnoreCase);
                                        text = Regex.Replace(text, @"( +)", " ");

                                        int snippetEnd = 50 < text.Length ? 50 : text.Length;

                                        log.Info("Extracted text! First few rows here.. \r\n " + text.Substring(0,snippetEnd));

                                        List<string> sentences = new List<string>();
                                        var RegEx_SentenceDelimiter = new Regex(@"(\.|\!|\?)");
                                        sentences = RegEx_SentenceDelimiter.Split(text).ToList();

                                        List<string> finalizedSentences = new List<string>();

                                        string sentenceCandidate = "";
                                        foreach (var sentence in sentences)
                                        {
                                            // sanitize
                                            if (sentence.Length < 5) continue;

                                            if (sentenceCandidate.Length + sentence.Length > 5120)
                                            {
                                                finalizedSentences.Add(sentenceCandidate);
                                                sentenceCandidate = sentence;
                                            }
                                            else
                                            {
                                                sentenceCandidate += " " + sentence;
                                            }
                                        }

                                        var analyzable = new List<MultiLanguageInput>();

                                        int i = 0;
                                        foreach (var s in finalizedSentences)
                                        {
                                            if (s.Length > 10) analyzable.Add(new MultiLanguageInput("en", i + "", s));
                                            i++;
                                        }

                                        if (keyPhrases.Count <= 0) RunTextAnalysis(ref keyPhrases, analyzable, log);
                                    }

                                    log.Info("All found key phrases were: ");
                                    foreach (var kp in keyPhrases)
                                    {
                                        log.Info(kp);
                                    }

                                    // then write the most important keyphrases back
                                    string description = "";
                                    foreach (var s in keyPhrases.Take(10))
                                    {
                                        description += s + "\r\n";
                                    }

                                    try
                                    {
                                        TextInfo ti = new CultureInfo("en-US", false).TextInfo;

                                        li["Title"] = ti.ToTitleCase(file.Name);
                                        li["Description0"] = ti.ToTitleCase(description.ToLower());

                                        li.Update();
                                    }
                                    catch (Exception ex)
                                    {
                                        log.Error(ex.Message);
                                    }

                                    try
                                    {
                                        ctx.Load(list.Fields);
                                        ctx.ExecuteQuery();

                                        var fieldnames = new string[] { "Keywords" };
                                        var field = list.GetFields(fieldnames).First();

                                        // setting managed metadata
                                        UpdateTaxonomyField(keyPhrases.Take(10).ToArray(), log, ctx, li, field);

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
                                    return req.CreateResponse(HttpStatusCode.InternalServerError, ex.Message);
                                }
                            }
                        }
                        else
                        {
                            return req.CreateResponse(HttpStatusCode.OK, "File was not pdf");
                        }

                        return req.CreateResponse(HttpStatusCode.OK, list.Id);
                    }
                    else
                    {
                        log.Info("List is not available on the site");
                        return req.CreateResponse(HttpStatusCode.OK, "List is not available on the site");
                    }
                }
            }
            catch (Exception ex)
            {
                log.Info("Error Message: " + ex.Message);
                message += "Error Message: " + ex.Message;
            }

            // Set name to query string or body data
            name = name ?? data?.name;

            log.Info("");

            return name == null
                ? req.CreateResponse(HttpStatusCode.BadRequest, "Please pass a name on the query string or in the request body")
                : req.CreateResponse(HttpStatusCode.OK, "Hello " + name + "\r\n\r\n" + message);
        }

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
                //messages += ex.Message;
                log.Warning(ex.Message);
            }
        }

        private static void UpdateTaxonomyField(string[] value, TraceWriter log, ClientContext ctx, ListItem item, Field field)
        {
            //var value = values[key];
            //if (value.GetType().IsArray)
            //{
                var taxSession = ctx.Site.GetTaxonomySession();
                var terms = new List<KeyValuePair<Guid, string>>();
                foreach (var arrayItem in value as object[])
                {
                    TaxonomyItem taxonomyItem;
                    Guid termGuid = Guid.Empty;

                    if (!Guid.TryParse(arrayItem as string, out termGuid))
                    {
                        // Assume it's a TermPath
                        taxonomyItem = ctx.Site.GetTaxonomyItemByPath(arrayItem as string);
                    }
                    else
                    {
                        taxonomyItem = taxSession.GetTerm(termGuid);
                        ctx.Load(taxonomyItem);
                        ctx.ExecuteQueryRetry();
                    }

                    

                    terms.Add(new KeyValuePair<Guid, string>(taxonomyItem.Id, taxonomyItem.Name));
                }

                TaxonomyField taxField = ctx.CastTo<TaxonomyField>(field);

                taxField.EnsureProperty(tf => tf.AllowMultipleValues);

                if (taxField.AllowMultipleValues)
                {
                    var termValuesString = String.Empty;
                    foreach (var term in terms)
                    {
                        termValuesString += "-1;#" + term.Value + "|" + term.Key.ToString("D") + ";#";
                    }

                    termValuesString = termValuesString.Substring(0, termValuesString.Length - 2);

                    var newTaxFieldValue = new TaxonomyFieldValueCollection(ctx, termValuesString, taxField);
                    taxField.SetFieldValueByValueCollection(item, newTaxFieldValue);

                    ctx.ExecuteQueryRetry();
                }
                else
                {
                    log.Info("You are trying to set multiple values in a single value field. Skipping values for field " );
                }
            //}
            //else
            //{
            //    Guid termGuid = Guid.Empty;
            //    if (!Guid.TryParse(value as string, out termGuid))
            //    {
            //        // Assume it's a TermPath
            //        var taxonomyItem = ClientContext.Site.GetTaxonomyItemByPath(value as string);
            //        termGuid = taxonomyItem.Id;
            //    }
            //    item[key as string] = termGuid.ToString();
            //}

                item.Update();
        }
    }
}
