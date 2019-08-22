using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Configuration;
using System.Data.SqlClient;
using System.Net;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using System.Text.RegularExpressions;
using System.Net.NetworkInformation;

namespace Personal_Assistant
{
    public partial class MainForm : Form
    {
        SpeechRecognitionEngine speechRecognitionEngine = null;
        SpeechSynthesizer Jarvis = new SpeechSynthesizer();
        // Adding some access modifier to declare a static member for list and events why we are using list,  to search, sort, and manipulate
        public static List<string> MsgList = new List<string>();
        public static List<string> MsgLink = new List<string>();
        // QEvent for check new emails
        public static String QEvent;
        //to count number of emails 
        int EmailNum = 0;
        //set global variable for username and password 
        string username;
        string password;
        string[] files, paths;
        //Variables for Weather 
        string title;
        string cdata;
        string temp;
        string condition;
        string high;
        string low;
        string humidity;
        string sunrise;
        string sunset;
        string windspeed;
        //=============//

        public MainForm()
        {
            InitializeComponent();
            try
            {
                // Set the language for speech engine
                speechRecognitionEngine = SetLanguage("en-US");
                //Event handler for recognized text 
                speechRecognitionEngine.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(Mainevent_SpeechRecognized);
                //This is for speak completed event
                Jarvis.SpeakCompleted += new EventHandler<SpeakCompletedEventArgs>(speak_completed);
                //Event for load grammar for speech engine 
                LoadGrammarAndCommands();
                // Using the system's default microphone
                speechRecognitionEngine.SetInputToDefaultAudioDevice();
                // Start listening 
                speechRecognitionEngine.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void speak_completed(object sender, SpeakCompletedEventArgs e)
        {
            ReadBtn.Enabled = true;
        }

        private void LoadGrammarAndCommands()
        {
            try
            {
                string connectionstring = ConfigurationManager.ConnectionStrings["MyDatabase"].ConnectionString;
                SqlConnection con = new SqlConnection(connectionstring);
                con.Open();
                SqlCommand sc = new SqlCommand();
                sc.Connection = con;
                sc.CommandText = "SELECT * FROM DefaultTable";
                //sc.CommandType = CommandType.TableDirect;
                SqlDataReader sdr = sc.ExecuteReader();
                while (sdr.Read())
                {
                    var Loadcmd = sdr["DefaultCommands"].ToString();
                    Grammar commandgrammar = new Grammar(new GrammarBuilder(new Choices(Loadcmd)));
                    speechRecognitionEngine.LoadGrammarAsync(commandgrammar);
                    
                }
                sdr.Close();
                con.Close();
            }
            catch (Exception ex)
            {
                Jarvis.SpeakAsync("I've detected an in valid entry in your web commands, possibly a blank line. web commands will case to work until it is fixed." + ex.Message);
            }
        }
        private void LoadGmailInfo()
        {
            try
            {
                string connectionstring = ConfigurationManager.ConnectionStrings["MyDatabase"].ConnectionString;
                SqlConnection con = new SqlConnection(connectionstring);
                con.Open();
                SqlCommand sc = new SqlCommand();
                sc.Connection = con;
                sc.CommandText = "SELECT * FROM GmailInfo";
                //sc.CommandType = CommandType.TableDirect;
                SqlDataReader sdr = sc.ExecuteReader();
                while (sdr.Read())
                {
                    username = sdr["Email"].ToString();
                    password = sdr["Password"].ToString();
                }
                sdr.Close();
                con.Close();
            }
            catch (Exception ex)
            {
                Jarvis.SpeakAsync("I've detected an in valid entry in your web commands, possibly a blank line. web commands will cease to work until it is fixed." + ex.Message);
            }
        }
        private void Mainevent_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //System Username 
            string Name = Environment.UserName;
            //Recognized Spoken words result is e.Result.Text
            string speech = e.Result.Text;
            //Debug_Livetxt.Text += "You said : " + e.Result.Text + "\n";
            //Switch to e.Result.Text
            switch (speech)
            {
                //Greetings
                case "hello":
                    Jarvis.SpeakAsync("hello " + Name);
                    System.DateTime timenow = System.DateTime.Now;
                    if (timenow.Hour >= 5 && timenow.Hour < 12)
                    { Jarvis.SpeakAsync("Goodmorning " + Name); }
                    if (timenow.Hour >= 12 && timenow.Hour < 18)
                    { Jarvis.SpeakAsync("Good afternoon " + Name); }
                    if (timenow.Hour >= 18 && timenow.Hour < 24)
                    { Jarvis.SpeakAsync("Good evening " + Name); }
                    if (timenow.Hour < 5)
                    { Jarvis.SpeakAsync("Hello " + Name + ", you are still awake you should go to sleep, it's getting late"); }
                    break;
                case "what time is it":
                    System.DateTime now = System.DateTime.Now;
                    string time = now.GetDateTimeFormats('t')[0];
                    Jarvis.SpeakAsync(time);
                    break;
                case "what day is it":
                    string day = "Today is," + System.DateTime.Now.ToString("dddd");
                    Jarvis.SpeakAsync(day);
                    break;
                case "what is the date":
                case "what is todays date":
                    string date = "The date is, " + System.DateTime.Now.ToString("dd MMM");
                    Jarvis.SpeakAsync(date);
                    date = "" + System.DateTime.Today.ToString(" yyyy");
                    Jarvis.SpeakAsync(date);
                    break;
                case "who are you":
                    Jarvis.SpeakAsync("i am your personal assistant");
                    Jarvis.SpeakAsync("i can read email, weather report, i can search web for you, anything that you need like a personal assistant do, you can ask me question i will reply to you");
                    break;
                case "what is my name":
                    Jarvis.SpeakAsync(Name);
                    break;
                case "get all emails":
                case "get all inbox emails":
                    EmailBtn.PerformClick();
                    AllEmails();
                    break;
                case "check for new emails":
                    EmailBtn.PerformClick();
                    QEvent = "Checkfornewemails";
                    Jarvis.SpeakAsyncCancelAll();
                    EmailNum = 0;
                    CheckEmails();
                    break;
                case "read the email":
                    Jarvis.SpeakAsyncCancelAll();
                    try
                    {
                        Jarvis.SpeakAsync(MsgList[EmailNum]);
                    }
                    catch
                    {
                        Jarvis.SpeakAsync("There are no emails to read");
                    }
                    break;
                case "next email":
                    Jarvis.SpeakAsyncCancelAll();
                    try
                    {
                        EmailNum += 1;
                        Jarvis.SpeakAsync(MsgList[EmailNum]);
                    }
                    catch
                    {
                        EmailNum -= 1;
                        Jarvis.SpeakAsync("There are no further emails");
                    }
                    break;
                case "previous email":
                    Jarvis.SpeakAsyncCancelAll();
                    try
                    {
                        EmailNum -= 1;
                        Jarvis.SpeakAsync(MsgList[EmailNum]);
                    }
                    catch
                    {
                        EmailNum += 1;
                        Jarvis.SpeakAsync("There are no previous emails");
                    }
                    break;
                //This is for text reader
                case "start reading":
                    //ReadBtn.PerformClick();
                    if (tabControl1.SelectedIndex == 1)
                    {
                        Jarvis.SpeakAsync(Readtxt.Text);
                    }
                    if (tabControl1.SelectedIndex == 3)
                    {
                        Jarvis.SpeakAsync(convertedtxt.Text);
                    }
                    break;
                case "pause":
                    //PauseBtn.PerformClick();
                    if (tabControl1.SelectedIndex == 1)
                    {
                        if(Jarvis.State == SynthesizerState.Speaking)
                        Jarvis.Pause();
                    }
                    if (tabControl1.SelectedIndex == 3)
                    {
                        if (Jarvis.State == SynthesizerState.Speaking)
                            Jarvis.Pause();
                    }
                    break;
                case "resume":
                    PauseBtn.PerformClick();
                    if (tabControl1.SelectedIndex == 1)
                    {
                        if (Jarvis.State == SynthesizerState.Speaking)
                            Jarvis.Resume();
                    }
                    if (tabControl1.SelectedIndex == 3)
                    {
                        if (Jarvis.State == SynthesizerState.Speaking)
                            Jarvis.Resume();
                    }
                    break;
                case "stop":
                    StopBtn.PerformClick();
                    break;
                case "open text file":
                    Open_FileBtn.PerformClick();
                    break;
                //---untill here--- //
                case "change voice speed to minis two":
                    Jarvis.Rate = -2;
                    break;
                case "change voice speed to minis four":
                    Jarvis.Rate = -4;
                    break;
                case "change voice speed to minis six":
                    Jarvis.Rate = -6;
                    break;
                case "change voice speed to minis eight":
                    Jarvis.Rate = -8;
                    break;
                case "change voice speed to minis ten":
                    Jarvis.Rate = -10;
                    break;
                case "change voice speed back to normal":
                    Jarvis.Rate = 0;
                    break;
                case "change voice speed to two":
                    Jarvis.Rate = 2;
                    break;
                case "change voice speed to four":
                    Jarvis.Rate = 4;
                    break;
                case "change voice speed to six":
                    Jarvis.Rate = 6;
                    break;
                case "change voice speed to eight":
                    Jarvis.Rate = 8;
                    break;
                case "change voice speed to ten":
                    Jarvis.Rate = 10;
                    break;
                    //For Media Player 
                case "i want to add music":
                case "i want to add video":
                    Jarvis.Speak("choose, music file from your drives");
                    Add_Music.PerformClick();
                    break;
                case "play music":
                case "play video":
                    PlayBtn.PerformClick();
                    break;
                case "stop media player":
                    MediaStopBtn.PerformClick();
                    break;
                case "fast forward":
                    FastfarwardBtn.PerformClick();
                    break;
                case "fast reverse":
                    FastReverseBtn.PerformClick();
                    break;
                case "media player resume":
                    PlayBtn.PerformClick();
                    break;
                case "media player pause":
                    PlayBtn.PerformClick();
                    break;
                case "media player previous":
                    PreviousBtn.PerformClick();
                    break;
                case "media player next":
                    NextBtn.PerformClick();
                    break;
                case "activate full screen mode":
                    FullScreen.PerformClick();
                    break;
                case "exit full screen":
                    FullScreen.PerformClick();
                    break;
                case "mute volume":
                case "volume down":
                    Unmute_Volum.PerformClick();
                    break;
                case "unmute volume":
                case "volume up":
                    Unmute_Volum.PerformClick();
                    break;
                //This is for news reader
                case "get bing news":
                    GetBingNews();
                    break;
                //Untill here // 
                //Weather grammar 
                case "get weather report":
                    Jarvis.SpeakAsync("ok, " + Name + " here is the weather report");
                    GetWeather();
                    break;
            }
        }
        public void GetWeather()
        {
            webBrowser2.Navigate("https://www.yahoo.com/news/weather/");
            //Converts the value of objects to strings based on the formats specified and inserts them into another string
            String query = String.Format("https://query.yahooapis.com/v1/public/yql?q=select * from weather.forecast where woeid in (select woeid from geo.places(1) where text='city, state')&format=xml&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys");
            //than we going to replace city with givan in database table city name or textbox
            string output = query.Replace("city", "attica");
            //Represents an XML document. You can use this class to load, validate, edit, add, and position XML in a document.
            XmlDocument wData = new XmlDocument();
            //Loads the XML document from the specified XmlReader.
            wData.Load(output);
            //Resolves, adds, and removes namespaces to a collection and provides scope management for these namespaces. 
            XmlNamespaceManager manager = new XmlNamespaceManager(wData.NameTable);
            //Adds the given namespace to the collection
            manager.AddNamespace("yweather", "http://xml.weather.yahoo.com/ns/rss/1.0");
            //Represents a single node in the XML document Which is query results and channel in yahoo api
            XmlNode channel = wData.SelectSingleNode("query").SelectSingleNode("results").SelectSingleNode("channel");
            //Represents an ordered collection of nodes
            XmlNodeList nodes = wData.SelectNodes("query/results/channel");

            try
            {
                temp = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", manager).Attributes["temp"].Value;

                condition = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", manager).Attributes["text"].Value;

                high = channel.SelectSingleNode("item").SelectSingleNode("yweather:forecast", manager).Attributes["high"].Value;

                low = channel.SelectSingleNode("item").SelectSingleNode("yweather:forecast", manager).Attributes["low"].Value;

                humidity = channel.SelectSingleNode("yweather:atmosphere", manager).Attributes["humidity"].Value;

                windspeed = channel.SelectSingleNode("yweather:wind", manager).Attributes["chill"].Value;

                sunrise = channel.SelectSingleNode("yweather:astronomy", manager).Attributes["sunrise"].Value;

                sunset = channel.SelectSingleNode("yweather:astronomy", manager).Attributes["sunset"].Value;

                cdata = channel.SelectSingleNode("item").SelectSingleNode("description").Value;

                temptxt.Text = "The temperature is : " + temp;
                Jarvis.SpeakAsync(temptxt.Text);
                conditiontxt.Text = "The condition is :" + condition;
                Jarvis.SpeakAsync(conditiontxt.Text);
                hightxt.Text = "The high is :" + high;
                Jarvis.SpeakAsync(hightxt.Text);
                lowtxt.Text = "The low is :" + low;
                Jarvis.SpeakAsync(lowtxt.Text);
                humiditytxt.Text = "The humidity is :" + humidity;
                Jarvis.SpeakAsync(humiditytxt.Text);
                windspeedtxt.Text = "wind speed is :" + windspeed;
                Jarvis.SpeakAsync(windspeedtxt.Text + " miles per hour");
                sunrisetxt.Text = "sun rise at :" + sunrise;
                Jarvis.SpeakAsync(sunrisetxt.Text);
                sunsetxt.Text = "sun set at :" + sunset;
                Jarvis.SpeakAsync(sunsetxt.Text);
            }

            catch(Exception ex)
            {
                MessageBox.Show("Error Reciving data", ex.Message);
            }
        }
        public void GetBingNews()
        {
            string checkinternet = NetworkInterface.GetIsNetworkAvailable().ToString();
            if (checkinternet != "True")
            {
                Jarvis.SpeakAsync("Please check your internet connection, before the news broadcast panel, work properly");
            }
            else
            {
                Jarvis.SpeakAsync("todays latest news is");
                convertedtxt.Clear();
                //it is common methods for sending data to and receiving data from a resource identified by a URI.
                WebClient webclient = new WebClient();
                // Downloads the requested resource as a String. The resource to download is specified as a String containing the URI.
                string page = webclient.DownloadString("https://www.bing.com/news/search?q=World&nvaug=%5bNewsVertical+Category%3d%22rt_World%22%5d&FORM=NSBABR");
                webBrowser1.Navigate("https://www.bing.com/news/search?q=World&nvaug=%5bNewsVertical+Category%3d%22rt_World%22%5d&FORM=NSBABR");
                //than we parse the html div tag and we will store in string variable
                string news = "<div class=\"snippet\">(.*?)</div>";
                //Searches the specified input string for the first occurrence of the regular expression specified in the Regex constructor.
                foreach (Match match in Regex.Matches(page, news))
                {
                    //Gets a collection of groups matched by the regular expression
                    convertedtxt.Text += match.Groups[1].Value;
                }
                ReadBtn.PerformClick();
            }
        }
        private SpeechRecognitionEngine SetLanguage(string preferredCulture)
        {
            //Checking for installed language and comparing with our given parameter preferredCulture to set speech recognition engine language
            foreach (RecognizerInfo config in SpeechRecognitionEngine.InstalledRecognizers())
            {
                if (config.Culture.ToString() == preferredCulture)
                {
                    speechRecognitionEngine = new SpeechRecognitionEngine(config);
                    break;
                }
            }

            // if the desired culture is not found, then load default
            if (speechRecognitionEngine == null)
            {
                MessageBox.Show("The desired languages is not installed on this machine, the speech-engine will continue using "
                    + SpeechRecognitionEngine.InstalledRecognizers()[0].Culture.ToString() + " as the default language.",
                    "Culture " + preferredCulture + " not found!");
                speechRecognitionEngine = new SpeechRecognitionEngine(SpeechRecognitionEngine.InstalledRecognizers()[0]);
            }

            return speechRecognitionEngine;
        }

        private void LeftSideMenuBtn_Click(object sender, EventArgs e)
        {
            if(LeftSideMenu.Width == 60)
            {
                LeftSideMenu.Width = 260;
                Title_lbl.Text = "Personal Assistant";
            }
            else
            {
                LeftSideMenu.Width = 60;
                Title_lbl.Text = "PA";
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            LeftSideMenu.Width = 60;
            RightSideMenu.Width = 0;
            Title_lbl.Text = "PA";
            Debug_Livetxt.SelectionIndent += 15;
            Debug_Livetxt.SelectionRightIndent += 20;
            MediaPlayer.uiMode = "none";
            Debug_Livetxt.Text += "You :" + "Do you know how to make padding in rich text box" + "\n"+"\r";
            Debug_Livetxt.Text += "J.A.R.V.I.S : " + "Is it possible to add padding into a Rich Text Box control between the text and the border? I tried docking a rich text box inside of a panel, with its padding for all four side set to 10 and that accomplished what I wanted.Except once the vertical scroll bar for the rich text box is needed that gets padded as well.";
        }
        private void HideTabPages()
        {
            //If you want to hide TabPages of TabControl
            tabControl1.Appearance = TabAppearance.FlatButtons;
            tabControl1.ItemSize = new Size(0, 1);
            tabControl1.SizeMode = TabSizeMode.Fixed;
        }
        private void RightMenuBtn_Click(object sender, EventArgs e)
        { 
            if(RightSideMenu.Width == 0)
            {
                RightSideMenu.Width = 260;
            }
            else
            {
                RightSideMenu.Width = 0;
            }
        }
        private void AllEmails()
        {
            try
            {
                WebClient objClient = new WebClient();
                string response;
                string title;
                string tagline;
                string summary;

                //Creating a new xml document
                XmlDocument doc = new XmlDocument();

                //Logging in Gmail server to get data
                objClient.Credentials = new System.Net.NetworkCredential(username, password);
                //reading data and converting to string
                response = Encoding.UTF8.GetString(
                           objClient.DownloadData(@"https://mail.google.com/mail/feed/atom"));

                response = response.Replace(
                     @"<feed version=""0.3"" xmlns=""http://purl.org/atom/ns#"">", @"<feed>");

                //loading into an XML so we can get information easily
                doc.LoadXml(response);
                //Counting all emails 
                string total_mails = doc.SelectSingleNode(@"/feed/fullcount").InnerText;
                //display amount of email with label 
                total_emails.Text = total_mails;

                Jarvis.SpeakAsync("Total numbers of emails are, " + total_mails + "email is exist in gmail inbox");
                //this is display tag line in to textbox 
                tagline = doc.SelectSingleNode("/feed/tagline").InnerText;
                Email_Tags_Lines.Text = tagline;
                //this will read the email tag lines
                Jarvis.SpeakAsync("sir, you have " + tagline);

                //Reading the title and the summary for every email
                foreach (XmlNode node in doc.SelectNodes(@"/feed/entry"))
                {
                    //Reading the email author from atom feed 
                    Email_Author.Text = node.SelectSingleNode("author").SelectSingleNode("name").InnerText;

                    MSGFrom.Rows.Add(node.SelectSingleNode("author").SelectSingleNode("name").InnerText);
                    Jarvis.SpeakAsync("Email from, " + Email_Author.Text.ToString());
                    //Reading a title of email 
                    title = node.SelectSingleNode("title").InnerText;
                    Message_title.Text = title;
                    Jarvis.SpeakAsync("Sir, mail is about, " + title.ToString());
                    //GETTING email summery
                    summary = node.SelectSingleNode("summary").InnerText;
                    Message_Summery.Text = summary.ToString();
                    Jarvis.SpeakAsync("And the summary is, " + summary.ToString());
                }
            }
            catch (Exception ex)
            {
                Jarvis.SpeakAsync("Please login to your gmail account and turn on less secure apps before this get work" + ex.Message);
                MessageBox.Show("Login to your gmail account and turn on less secure apps before this get work", ex.Message);
                System.Diagnostics.Process.Start("https://support.google.com/accounts/answer/6010255?hl=en");
            }
        }
        private void CheckEmails()
        {
            string GmailAtomUrl = "https://mail.google.com/mail/feed/atom";
            //we are going to resolves external XML resources for credentials 
            XmlUrlResolver xmlResolver = new XmlUrlResolver();
            xmlResolver.Credentials = new NetworkCredential(username, password);
            //XmlTextReader fast access to XML data from gmail atom url
            XmlTextReader xmlReader = new XmlTextReader(GmailAtomUrl);
            xmlReader.XmlResolver = xmlResolver;
            try
            {
                //Gets the Uniform Resource Identifier for example from (http://purl.org/atom/ns#) 
                XNamespace ns = XNamespace.Get("http://purl.org/atom/ns#");
                //Initializes a new instance of the XDocument class to load uniform resource identifier from google feed atom
                XDocument xmlFeed = XDocument.Load(xmlReader);


                var emailItems = from item in xmlFeed.Descendants(ns + "entry")
                                 select new
                                 {
                                     Author = item.Element(ns + "author").Element(ns + "name").Value,
                                     Title = item.Element(ns + "title").Value,
                                     Link = item.Element(ns + "link").Attribute("href").Value,
                                     Summary = item.Element(ns + "summary").Value
                                 };
                MainForm.MsgList.Clear();
                MainForm.MsgLink.Clear();

                foreach (var item in emailItems)
                {
                    if (item.Title == String.Empty)
                    {
                        MainForm.MsgList.Add("Message from " + item.Author + ", There is no subject and the summary reads, " + item.Summary);
                        MainForm.MsgLink.Add(item.Link);
                    }
                    else
                    {
                        MainForm.MsgList.Add("Message from " + item.Author + ", The subject is " + item.Title + " and the summary reads, " + item.Summary);
                        MainForm.MsgLink.Add(item.Link);
                    }
                }

                if (emailItems.Count() > 0)
                {
                    if (emailItems.Count() == 1)
                    {
                        Jarvis.SpeakAsync("You have 1 new email");
                    }
                    else
                    {
                        Jarvis.SpeakAsync("You have " + emailItems.Count() + " new emails");
                    }
                }
                else if (MainForm.QEvent == "Checkfornewemails" && emailItems.Count() == 0)
                {
                    Jarvis.SpeakAsync("You have no new emails");
                    MainForm.QEvent = String.Empty;
                }
            }
            catch(Exception ex)
            {
                Jarvis.SpeakAsync("You have submitted invalid log in information");
                Jarvis.SpeakAsync("Please login to your gmail account and turn on less secure apps before this get work" + ex.Message);
            }
        }

        private void EmailBtn_Click(object sender, EventArgs e)
        {
            LoadGmailInfo();
            tabControl1.SelectedIndex = 0;
        }

        private void Open_FileBtn_Click(object sender, EventArgs e)
        {
            if(Jarvis.State == SynthesizerState.Speaking)
                Jarvis.SpeakAsyncCancelAll();
            Readtxt.Clear();
            Jarvis.SpeakAsync("choose a text file from your, drives");

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            //Filter for type of the file we are going to open
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|rtf files (*.rtf)|*.rtf|All files (*.*)|*.*";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    //FileNames Gets the file names of all selected files in the dialog box
                    string strfilename = openFileDialog1.FileName;
                    //Here we are using System.IO class to reads the lines of a file
                    string filetext = File.ReadAllText(strfilename);
                    //than we are going to pass the filetxt to over textbox which is Readtxt.Text
                    Readtxt.Text = filetext;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        private void ReadBtn_Click(object sender, EventArgs e)
        {
            if(Jarvis.State == SynthesizerState.Speaking)
                Jarvis.SpeakAsyncCancelAll();
            ReadBtn.Enabled = false;
            PauseBtn.Enabled = true;
            if(tabControl1.SelectedIndex == 1)
            {
                Jarvis.SpeakAsync(Readtxt.Text);
            }
            if (tabControl1.SelectedIndex == 3)
            {
                Jarvis.SpeakAsync(convertedtxt.Text);
            }
        }

        private void PauseBtn_Click(object sender, EventArgs e)
        {
            if (Jarvis.State == SynthesizerState.Speaking)
            {
                Jarvis.Pause();
                PauseBtn.Text = "Resume";
            }
            else 
            {
                Jarvis.Resume();
                PauseBtn.Text = "Pause";
            }
        }

        private void StopBtn_Click(object sender, EventArgs e)
        {
            if(Jarvis.State == SynthesizerState.Speaking)
            Jarvis.SpeakAsyncCancelAll();
            ReadBtn.Enabled = true;
        }

        private void PlayBtn_Click(object sender, EventArgs e)
        {
            if(MediaPlayer.playState == WMPLib.WMPPlayState.wmppsPlaying)
            {
                PlayBtn.BackgroundImage = Properties.Resources.media_btnpause;
                MediaPlayer.Ctlcontrols.pause();
            }
            else if (MediaPlayer.playState == WMPLib.WMPPlayState.wmppsPaused)
            {
                PlayBtn.BackgroundImage = Properties.Resources.media_btnplay;
                MediaPlayer.Ctlcontrols.play();
            }
            else if (PlayList.SelectedIndex > 0)
            {
                MediaPlayer.URL = paths[PlayList.SelectedIndex];
                MediaPlayer.Ctlcontrols.play();
            }
            else
            {
                PlayList.SelectedIndex = 0;
                MediaPlayer.Ctlcontrols.play();
            }
        }

        private void MediaStopBtn_Click(object sender, EventArgs e)
        {
            if (MediaPlayer.playState == WMPLib.WMPPlayState.wmppsPlaying)
            {
                MediaPlayer.Ctlcontrols.stop();
            }
            else
            {
                MediaPlayer.Ctlcontrols.play();
            }
        }

        private void PreviousBtn_Click(object sender, EventArgs e)
        {
            if (MediaPlayer.playState == WMPLib.WMPPlayState.wmppsPlaying)
            {
                if (PlayList.SelectedIndex > 0)
                {
                    MediaPlayer.Ctlcontrols.previous();
                    PlayList.SelectedIndex -= 1;
                    PlayList.Update();
                }
                else
                {
                    PlayList.SelectedIndex = 0;
                    PlayList.Update();
                }
            }
        }

        private void NextBtn_Click(object sender, EventArgs e)
        {
            if (MediaPlayer.playState == WMPLib.WMPPlayState.wmppsPlaying)
            {
                if (PlayList.SelectedIndex < (PlayList.Items.Count - 1))
                {
                    MediaPlayer.Ctlcontrols.next();
                    PlayList.SelectedIndex += 1;
                    PlayList.Update();
                }
                else
                {
                    PlayList.SelectedIndex = 0;
                    PlayList.Update();
                }
            }
        }

        private void FastReverseBtn_Click(object sender, EventArgs e)
        {
            MediaPlayer.Ctlcontrols.fastReverse();
        }

        private void FastfarwardBtn_Click(object sender, EventArgs e)
        {
            MediaPlayer.Ctlcontrols.fastForward();
        }

        private void Unmute_Volum_Click(object sender, EventArgs e)
        {
            if (MediaPlayer.settings.volume == 100)
            {
                MediaPlayer.settings.volume = 0;
                Unmute_Volum.BackgroundImage = Properties.Resources.mediaplayervolumedown;
            }
            else
            {
                MediaPlayer.settings.volume = 100;
                Unmute_Volum.BackgroundImage = Properties.Resources.mediaplayervolumeupp;
            }
        }

        private void PlayBackTimmer_Tick(object sender, EventArgs e)
        {
            PlayBackTimmer.Start();
            if (PlayList.SelectedIndex < files.Length - 1)
            {
                PlayList.SelectedIndex++;
                PlayBackTimmer.Enabled = false;
            }
            else
            {
                PlayList.SelectedIndex = 0;
                PlayBackTimmer.Enabled = false;
            }
            PlayBackTimmer.Stop();
        }

        private void FullScreen_Click(object sender, EventArgs e)
        {
            if (MediaPlayer.playState == WMPLib.WMPPlayState.wmppsPlaying)
            {
                MediaPlayer.fullScreen = true;
            }
            else
            {
                MediaPlayer.fullScreen = false;
            }
        }

        private void Volum_Speed_Scroll(object sender, EventArgs e)
        {
            MediaPlayer.settings.volume = Volum_Speed.Value;
        }

        private void PlayList_SelectedIndexChanged(object sender, EventArgs e)
        {
            MediaPlayer.URL = paths[PlayList.SelectedIndex];
        }

        private void News_Readbtn_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 3;
            Jarvis.SpeakAsync(convertedtxt.Text);
        }

        private void News_PauseBtn_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 3;
            if (Jarvis.State == SynthesizerState.Speaking)
            {
                Jarvis.Pause();
                News_PauseBtn.Text = "Resume";
            }
            else
            {
                Jarvis.Resume();
                News_PauseBtn.Text = "Pause";
            }
        }

        private void News_StopBtn_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 3;
            if (Jarvis.State == SynthesizerState.Speaking)
                Jarvis.SpeakAsyncCancelAll();
        }

        private void MinimizeBtn_Click(object sender, EventArgs e)
        {
            FormBorderStyle = FormBorderStyle.None;
            WindowState = FormWindowState.Minimized;
            TopMost = false;
        }

        private void MaximizeBtn_Click(object sender, EventArgs e)
        {
            FormBorderStyle = FormBorderStyle.None;
            WindowState = FormWindowState.Normal;
            TopMost = true;
        }

        private void CloseBtn_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void FullScreenBtn_Click(object sender, EventArgs e)
        {
            FormBorderStyle = FormBorderStyle.None;
            WindowState = FormWindowState.Maximized;
            TopMost = true;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 1;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 2;
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 3;
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 4;
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 5;
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 6;
        }

        private void SettingsBtn_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 7;
        }

        private void Add_Music_Click(object sender, EventArgs e)
        {
            string userName = System.Environment.UserName;

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = @"C:\Users\" + userName + "\\Documents\\MyMusic";
            ofd.Filter = "(mp3,wav,mp4,mov,wmv,mpg,avi,3gp,flv)|*.mp3;*.wav;*.mp4;*.3gp;*.avi;*.mov;*.flv;*.wmv;*.mpg|all files|*.*";
            ofd.Multiselect = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //SafeFileNames Gets the file name and extension for the file selected in the dialog box.The file name does not include the path.
                files = ofd.SafeFileNames;
                //FileNames Gets the file names of all selected files in the dialog box
                paths = ofd.FileNames;
                for (int i = 0; i < files.Length; i++)
                {
                    PlayList.Items.Add(files[i]);
                }
            }
        }
    }
}
