using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using System.Net;
using System.Collections.Specialized;
using Newtonsoft.Json.Linq;
using System.IO;
using Newtonsoft.Json;
using System.Threading;

using System.Configuration;

//using System.Web.Helpers;

using TikaOnDotNet.TextExtraction;
using System.Text.RegularExpressions;
using System.ComponentModel;

using Microsoft.Azure.CognitiveServices.Language.TextAnalytics;
using Microsoft.Azure.CognitiveServices.Language.TextAnalytics.Models;

namespace Koskila.KeywordManager
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private string url_topics = "https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/keyPhrases";
        private string path_extractableFile = System.Configuration.ConfigurationManager.AppSettings[""];

        private string key1 = System.Configuration.ConfigurationManager.AppSettings[""];

        private int maxSentenceLength = 25000;

        bool useWastefulLogic = false;

        public MainWindow()
        {
            InitializeComponent();
            InitTestData();

            // pause the main thread for a while to stop from getting 429 errors
            Thread.Sleep(1000);
        }

        private void InitTestData()
        {
            int[] values = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20 };

            foreach (var v in values)
            {
                ddl_ResultsCount.Items.Add(v);
            }
            ddl_ResultsCount.SelectedIndex = ddl_ResultsCount.Items.Count - 1;

            txtBox_pathToFiles.Text = path_extractableFile;

            //ICollectionView view = CollectionViewSource.GetDefaultView(resultsArea);
            //view.GroupDescriptions.Add(new PropertyGroupDescription("Title"));
            //view.SortDescriptions.Add(new SortDescription("Title", ListSortDirection.Ascending));
            //view.SortDescriptions.Add(new SortDescription("score", ListSortDirection.Ascending));
        }

        private void button_OK_Click(object sender, RoutedEventArgs e)
        {
            var data = new MultiLanguageBatchInput();

            // Extracting text
            var di = new DirectoryInfo(txtBox_pathToFiles.Text);
            int index = 0;

            var RegEx_SentenceDelimiter = new Regex(@"(\.|\!|\?)");

            string fulltext = "";

            foreach (FileInfo fi in di.GetFiles())
            {
                string path = fi.FullName;
                string title = fi.Name;

                var extractor = new TikaOnDotNet.TextExtraction.TextExtractor();
                var extractionResult = extractor.Extract(path);
                string text = extractionResult.Text;

                text = Regex.Replace(text, @"[\r\n\t\f\v]", " ");
                text = Regex.Replace(text, @"[^a-z.,!?]", " ", RegexOptions.IgnoreCase);
                text = Regex.Replace(text, @"( +)", " ");

                var values = new JObject();

                JArray documents = new JArray();
                Topic topic = new Topic();

                int sentenceCount = RegEx_SentenceDelimiter.Split(text).Length;

                //int factor = 1;
                //if ((double) sentenceCount / 1000 <= 1) factor = 1;
                //else factor = (sentenceCount / 1000) + 1;

                List<string> sentences = new List<string>();

                //if (useWastefulLogic)
                //{
                //    if (sentenceCount < 100)
                //    {
                //        var splitFactor = (100 / sentenceCount) + 1;

                //        // splitFactor tells us, into how many pieces each sentence needs to be split
                //        foreach (var sentenceCandidate in RegEx_SentenceDelimiter.Split(text))
                //        {
                //            sentences.Add(sentenceCandidate);

                //            for (int j = 1; j <= splitFactor; j++)
                //            {
                //                sentences.Add(" ");
                //            }
                //        }
                //    }
                //    else if (100 < sentenceCount && sentenceCount < 1000)
                //    {
                sentences = RegEx_SentenceDelimiter.Split(text).ToList();
                //    }
                //    else // sentenceCount >= 1000
                //    {
                //        int counter = 1;
                //        string t = "";
                //        int docId = 1;

                //        sentences = RegEx_SentenceDelimiter.Split(text).ToList();

                //        foreach (string sentence in sentences)
                //        {
                //            if (counter <= factor)
                //            {
                //                t += sentence;

                //                counter++;
                //            }
                //            else
                //            {
                //                Document d = new Document();
                //                d.id = docId;
                //                d.text = t;
                //                topic.documents.Add(d);

                //                t = "";
                //                t += sentence;
                //                counter = 1;

                //                docId++;
                //            }
                //        }
                //    }
                //}
                //else
                //{
                //    sentences = RegEx_SentenceDelimiter.Split(text).ToList();

                //    int maxSentencesPerDocument = sentences.Count / 100;

                //    int counter = 1;
                //    string t = "";
                //    int docId = 1;

                //    foreach (string sentence in sentences)
                //    {
                //        if ((t + ". " + sentence).Length > maxSentenceLength || counter >= maxSentencesPerDocument)
                //        {
                //            Document d = new Document();
                //            d.id = docId;
                //            d.text = t;
                //            topic.documents.Add(d);

                //            t = "";
                //            t += sentence;
                //            counter = 1;

                //            docId++;
                //        }
                //        else
                //        {
                //            t += ". " + sentence;

                //            counter++;
                //        }
                //    }
                //}

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
                //analyzable.Add(new MultiLanguageInput("en", 0 + "", fulltext));
                data.Documents = analyzable;



                //topic.stopWords.Add("world");

                //topic.stopPhrases.Add("world");


                //string result = "";

                ITextAnalyticsAPI client = new TextAnalyticsAPI();
                client.AzureRegion = AzureRegions.Westus;
                client.SubscriptionKey = key1;

                //JsonSerializerSettings jss = new JsonSerializerSettings();
                //jss.Formatting = Formatting.None;

                //string json = values.ToString();
                
                //json = JsonConvert.SerializeObject(topic, jss);

                try
                {
                    var result = client.KeyPhrases(data);

                    foreach (var row in result.Documents)
                    {
                        foreach (var kp in row.KeyPhrases)
                        {
                            AddMessage(kp);
                        }
                    }

                    //result = client.UploadString(url_topics, "POST", json);

                    //string requestId = client.ResponseHeaders.Get("operation-location");
                    //int topicCount = int.Parse(ddl_ResultsCount.SelectedItem.ToString());

                    //Thread thread = new Thread(delegate ()
                    //{
                    //    GetDataAndUpdate(requestId, title, index, topicCount);
                    //    // rest omitted for clarity
                    //});
                    //thread.IsBackground = true;
                    //thread.Start();

                    //// pause the main thread for a while to stop from getting 429 errors
                    //Thread.Sleep(60000);
                }
                catch (Exception ex)
                {
                    AddMessage(title + ": " + ex.Message);
                }   
            }
            index++;
            
        }

        private void GetDataAndUpdate(string requestId, string title, int wait, int topicCount)
        {
            // next try once a minute to fetch contents
            using (var client = new WebClient())
            {
                // setting correct headers
                client.Headers.Add("Ocp-Apim-Subscription-Key", key1);
                client.Headers.Add("Content-Type", "application/json");
                //client.Headers[HttpRequestHeader.ContentType] = "application/json";
                client.Headers.Add("Accept", "application/json");

                for (int i = 0; i < 15 + wait; i++)
                {
                    try
                    {
                        if (wait > i) throw new Exception(title + ": Need to wait a while.");
                        var responseJson = client.DownloadString(requestId);

                        var jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<TopicResponse>(responseJson);

                        switch (jsonData.status.ToLower())
                        {
                            case "running":
                                throw new Exception(title + ": Still running...");
                                break;
                            case "notstarted":
                                throw new Exception(title + ": Not started yet!");
                                break;
                            case "succeeded":
                                AddMessage(title + ": " + jsonData.status + " " + jsonData.message);

                                var sortedTopics = jsonData.operationProcessingResult.topics.ToList();
                                sortedTopics.Sort((p1, p2) => p1.score.CompareTo(p2.score));
                                sortedTopics.Reverse();
                                sortedTopics = sortedTopics.Take(topicCount).ToList();

                                foreach (TopicResponseTopic trt in sortedTopics)
                                {
                                    TopicResponseTopicSortingItem trtsi = new TopicResponseTopicSortingItem(trt);
                                    trtsi.title = title;
                                    AddData(title, trtsi);
                                }

                                break;
                            default:
                                throw new Exception(title + ": Don't know what happened.");
                                break;
                        }

                        break;
                    }
                    catch (Exception ex)
                    {
                        AddMessage(ex.Message);
                        Thread.Sleep(60000);
                    }
                }
            }
        }

        public void AddMessage(string text)
        {
            this.Dispatcher.Invoke(() =>
            {
                lst_Status.Items.Add(text);
            });
        }

        public void AddData(string text, TopicResponseTopicSortingItem trt)
        {
            this.Dispatcher.Invoke(() =>
            {
                var i = lst_Results.Items.Add(trt);
                //resultsArea.Items.GetItemAt(i)
            });
        }

        public void AddData(string text, string trt)
        {
            this.Dispatcher.Invoke(() =>
            {
                var i = lst_Results.Items.Add(trt);
                //resultsArea.Items.GetItemAt(i)
            });
        }
    }

    // SENDING DATA
    public class Topic
    {
        public List<Document> documents;
        public List<string> stopWords;
        public List<string> stopPhrases;

        public Topic () {
            this.documents = new List<Document>();
            this.stopWords = new List<string>();
            this.stopPhrases = new List<string>();
        }
    }

    public class Document
    {
        public int id;
        public string text;
    }

    // RECEIVING
    public class TopicResponse
    {
        public string status;
        public string createdDateTime;
        public string operationType;
        public string message;
        public TopicResponseOperationProcessingResult operationProcessingResult;
    }

    public class TopicResponseOperationProcessingResult
    {
        public List<TopicResponseTopic> topics;
        public List<TopicResponseTopicAssignment> topicAssignments;
        public List<TopicResponseError> errors;
        public string discriminator;
    }

    public class TopicResponseTopic
    {
        public string id;
        public double score;
        public string keyPhrase;
    }

    public class TopicResponseTopicAssignment
    {
        public int documentId;
        public string topicId;
        public double distance;
    }

    public class TopicResponseError
    {
        public int id;
        public string message;
    }


    // SORTING
    public class TopicResponseTopicSortingItem
    {
        public string id;
        public double score;
        public string keyPhrase;

        public TopicResponseTopicSortingItem(TopicResponseTopic trt)
        {
            id = trt.id;
            score = trt.score;
            keyPhrase = trt.keyPhrase;
        }

        public override string ToString()
        {
            return this.title + ": " + this.score + " " + this.keyPhrase;
        }

        public string title;
    }
}
