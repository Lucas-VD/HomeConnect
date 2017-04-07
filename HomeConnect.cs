using Microsoft.Speech.Recognition;
using Microsoft.Speech.Synthesis;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using WebSocketSharp;

namespace HomeConnect
{
    class Program
    {
        /*
         *  Created by Lucas-VD on 1st of April 2017
         *  This project is currently under a CC BY-NC-SA 4.0 license (Because I don't really know how licensing works)
         *  
         *  The project uses WebSocketSharp (https://github.com/sta/websocket-sharp)
         */

        // Initializing variables
        private static WebSocket ws;
        private static Dictionary<String, short> relays;
        private static SpeechSynthesizer ss = new SpeechSynthesizer();
        private static SpeechRecognitionEngine sre;
        private static short relayId;
        private static String command;

        // The path where the relay information is saved to
        private static readonly string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\relays.xml";
        // The address of the HomeCenter server
        private static readonly string serverIp = "{Insert your HomeCenter server IP-address here}";
        // The username and password for the HomeCenter server
        private static readonly string username = "{Insert the username of your HomeCenter server here}";
        private static readonly string password = "{Insert the password of your HomeCenter server here}";

        static void Main(string[] args)
        {
            XmlDocument doc = new XmlDocument();
            CookieCollection cookies = new CookieCollection();

            // Check if there is already a copy of the HomeCenter webpage
            // If there is we can choose to update it
            LogHelper.Info("Do you want to load a new web page? (y/n)");

            if (Console.ReadLine().Equals("y") || !File.Exists(path))
            {
                // Get the PHP cookies to retrieve the webpage
                cookies = GetCookies();

                // Retrieve the webpage itself
                SaveWebPage(cookies, doc);

                // Parse the relay information into a Dictionary<string, short> [name, id]
                LoadRelays(doc);

                // Save the relay dictionary for futur use
                File.WriteAllText(path, JsonConvert.SerializeObject(relays));

            }
            else
                relays = JsonConvert.DeserializeObject<Dictionary<String, short>>(File.ReadAllText(path)); // Load the previously saved relay information

            // Initialize the WebSocket to communicate to the HomeCenter server
            InitializeWebSocket();

            // Initialize the Microsoft Speech service
            InitializeSpeech(ss, sre);

            // Press a key to close down the application
            // TODO: fancy this up
            Console.ReadLine();
            ws.Close();
            Environment.Exit(0);
        }

        /// <summary>
        /// Initializes the WebSocket to the HomeCenter server
        /// </summary>
        private static void InitializeWebSocket()
        {
            // Set up the socket
            LogHelper.Info("WebSocket inititionization...");
            ws = new WebSocket("ws://" + serverIp + "/events");

            // Display a message to let the user know the connection succeeded
            ws.OnOpen += (sender, e) =>
            {
                LogHelper.Info("WebSocket opened!");
            };

            // Connect the WebSocket
            ws.Connect();
        }

        /// <summary>
        /// Initializes the Microsoft Speech system
        /// </summary>
        /// <param name="ss">SpeechSynthesizer object</param>
        /// <param name="sre">SpeechRecognitionEngine object</param>
        private static void InitializeSpeech(SpeechSynthesizer ss, SpeechRecognitionEngine sre)
        {
            sre = new SpeechRecognitionEngine(new System.Globalization.CultureInfo("nl-NL"));

            // Set up the input and output devices
            ss.SetOutputToDefaultAudioDevice();
            sre.SetInputToDefaultAudioDevice();

            ss.SelectVoice("Microsoft Server Speech Text to Speech Voice (nl-NL, Hanna)");

            // Add the speechrecogniser method
            sre.SpeechRecognized += SpeechRecognized;

            // Adds the two control words for turning on and off lights
            Choices onOffCmd = new Choices();
            onOffCmd.Add("aan");
            onOffCmd.Add("uit");

            // Adds the name of all of the relays
            Choices relaysCmd = new Choices();
            foreach (KeyValuePair<String, short> relay in relays)
                relaysCmd.Add(relay.Key);

            // Builds the command: [name relay] + [on/off]
            GrammarBuilder commandGb = new GrammarBuilder();
            commandGb.Append(relaysCmd);
            commandGb.Append(onOffCmd);

            Grammar commandGrammar = new Grammar(commandGb);

            // Adds yes and no to confirm turning on and off relays
            Choices yesNoCmd = new Choices();
            yesNoCmd.Add("ja");
            yesNoCmd.Add("nee");

            GrammarBuilder yesNoGb = new GrammarBuilder();
            yesNoGb.Append(yesNoCmd);
            Grammar yesNoGrammar = new Grammar(yesNoGb);

            // Loads the grammar
            sre.LoadGrammarAsync(commandGrammar);
            sre.LoadGrammarAsync(yesNoGrammar);
            sre.RecognizeAsync(RecognizeMode.Multiple);

            // Lets the user know the application is listening
            Speak(ss, "Ik ben klaar om te luisteren!");
        }

        /// <summary>
        /// Requests the cookies from the HomeCenter login page
        /// </summary>
        /// <returns>A CookieCollection which contains the PHP userId</returns>
        private static CookieCollection GetCookies()
        {
            LogHelper.Info("Getting cookies from login page...");
            // Create a WebRequest to get the cookies
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("http://" + serverIp + "/plug/index.php");
            CookieCollection cookies = new CookieCollection();
            request.CookieContainer = new CookieContainer();
            request.CookieContainer.Add(cookies);
            //Get the response from the server and save the cookies from the first request..
            try
            {
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                cookies = response.Cookies;
            }
            catch (Exception e)
            {
                LogHelper.Error("We encountered a problem while pinging the login page: " + e.Message);
            }

            // Check if we got the cookies
            if (cookies != null)
                LogHelper.Info("Cookies saved!");
            else
                LogHelper.Error("We encountered a problem while saving cookies!");

            return cookies;
        }

        /// <summary>
        /// Save the webpage
        /// </summary>
        /// <param name="cookies">The cookies which contain the PHP sessionid</param>
        /// <param name="doc">The XmlDocument to save the web page to</param>
        private static void SaveWebPage(CookieCollection cookies, XmlDocument doc)
        {
            LogHelper.Info("Logging in...");
            // Set up the WebRequest
            string data = "username=" + username + "&password=" + password + "&selectMode=tablet&screenwidth=1349&screenheight=637&Submit=Login";
            HttpWebRequest getRequest = (HttpWebRequest)WebRequest.Create("http://" + serverIp + "/plug/index.php?action=check");
            getRequest.CookieContainer = new CookieContainer();
            getRequest.CookieContainer.Add(cookies);
            getRequest.Method = WebRequestMethods.Http.Post;
            getRequest.UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.2 (KHTML, like Gecko) Chrome/15.0.874.121 Safari/535.2";
            getRequest.AllowWriteStreamBuffering = true;
            getRequest.ProtocolVersion = HttpVersion.Version11;
            getRequest.AllowAutoRedirect = true;
            getRequest.ContentType = "application/x-www-form-urlencoded";

            // Send the HTTP POST command to retrieve the web page
            byte[] byteArray = Encoding.ASCII.GetBytes(data);
            getRequest.ContentLength = byteArray.Length;
            using (Stream newStream = getRequest.GetRequestStream())
                newStream.Write(byteArray, 0, byteArray.Length);

            // Check is the response was successful
            HttpWebResponse getResponse = (HttpWebResponse)getRequest.GetResponse();
            if (getResponse.StatusCode.Equals(HttpStatusCode.OK))
                LogHelper.Info("Webpage request completed!");
            else
                LogHelper.Error("We encountered a problem while requesting the webpage!");

            // Parse the response into the given XmlDocument
            using (StreamReader sr = new StreamReader(getResponse.GetResponseStream()))
                doc.LoadXml(sr.ReadToEnd());
        }

        /// <summary>
        /// This method reads the XmlDocument and extracts the dimmers and relays with their name and ID
        /// </summary>
        /// <param name="doc">The XmlDocument which contains the HomeCenter controls</param>
        /// <returns>A Dictionary<string, short> containing all of the HomeCenter elements [element name, ID]</string></returns>
        private static Dictionary<string, short> LoadRelays(XmlDocument doc)
        {
            Dictionary<String, short> relays = new Dictionary<String, short>();

            LogHelper.Info("Loading relays into memory...");
            // First retrieve all of the "<a>" nodes
            foreach (XmlNode node in doc.OwnerDocument.SelectNodes("//a"))
            {
                // Check if the class name contains the word "id"
                if (node.Attributes["class"].Value.ToLower().Contains("id"))
                {
                    // Check if the element has a name (data-title) and id (data-name)
                    if (!node.Attributes["data-title"].Value.Equals("") && !node.Attributes["data-name"].Value.Equals(""))
                    {
                        // Try to parse the values into the dictionary
                        try
                        {
                            string type = Regex.Match(node.OuterXml, @" name=""\d+").Value.Substring(7);
                            if (type.Equals("0") || type.Equals("2"))
                                relays.Add(node.Attributes["data-title"].Value, Convert.ToInt16(node.Attributes["data-name"].Value));
                        }
                        catch (ArgumentException e)
                        {
                            // The HomeCenter server contains a lot of duplicates and these tend to throw errors
                            /*if (e.Message.Contains("sleutel") || e.Message.Contains("key"))
                                LogHelper.Info("DIT IS GEEN ERROR, maar er zijn duplicaten gevonden in de relays.");*/
                        }
                    }
                }
            }

            LogHelper.Info("Finished loading relays!");
            return relays;
        }

        /// <summary>
        /// Just speaks out the data and also logs it
        /// </summary>
        /// <param name="ss">The SpeechSynthesizer object used to "speak"</param>
        /// <param name="data">The string that needs to be said</param>
        private static void Speak(SpeechSynthesizer ss, string data)
        {
            LogHelper.Speak(data);
            ss.Speak(data);
        }

        /// <summary>
        /// Gets called when the application hears a command
        /// </summary>
        /// <param name="sender">Not used</param>
        /// <param name="e">The EventArguments of the call</param>
        private static void SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            // Get the command said and its confidence
            string txt = e.Result.Text;
            float confidence = e.Result.Confidence;
            LogHelper.Info("Recognized: " + txt);

            // Check if the recogniser was confident
            if (confidence < 0.60)
            {
                LogHelper.Warn("The command wasn't recognized!");
                return;
            }

            // Check if it was a confirmation command or a regular command
            if (txt.Equals("ja"))
            {
                SpeechConfirmation();
            }
            else if (txt.Equals("nee"))
            {
                relayId = -1;
                command = "";
                Speak(ss, "OK");
            }
            else if (relays.TryGetValue(txt.Substring(0, txt.Length - 4), out short id))
            {
                // Passes the speech command as well as the the id of the relay
                SpeechCommand(id, txt);
            }
            else
                LogHelper.Warn("The command wasn't recognized!");
        }

        /// <summary>
        /// Asks for a confirmation to turn the lights on or off
        /// </summary>
        /// <param name="id">The id of the relay</param>
        /// <param name="txt">The command said by the user</param>
        private static void SpeechCommand(short id, String txt)
        {
            relayId = id;
            command = txt;
            // Aks if the user really wants to control this relay
            Speak(ss, String.Format("Wil je de lichten {0} zeker {1}", command.Substring(0, command.Length - 4), command.Contains("aan") ? "aandoen?" : "uitdoen?"));
        }

        /// <summary>
        /// Confirms the command and sends a message to th server to control the relay
        /// </summary>
        private static void SpeechConfirmation()
        {
            // Check if the user really said he/she wants to control a light
            if (relayId != -1 && !command.IsNullOrEmpty())
            {
                // Send the message to the server
                SendMessageToServer(relayId, command.Contains("aan"));
                // Let the user know the message has been sent
                Speak(ss, String.Format("Ik doe {0} {1}.", command.Substring(0, command.Length - 4), command.Contains("aan") ? "aan" : "uit"));
            }
            else
                LogHelper.Error("The was no command found that had to be confirmed!");

            // Reset the id and command variables to avoid problems
            relayId = -1;
            command = "";
        }

        /// <summary>
        /// Forms a command to send to the HomeCenter server
        /// </summary>
        /// <param name="id">The relay id</param>
        /// <param name="turnOn">Do we need to switch the relay on or off?</param>
        private static void SendMessageToServer(short id, bool turnOn)
        {
            LogHelper.Info("Sending message to server with the id of " + id);
            short[] data = new short[3];

            // Forms the command based on the HomeCenter protocol [28, id, (0-255)]
            data[0] = 28;
            data[1] = id;
            data[2] = turnOn ? (short)255 : (short)0;

            // Create a packet out of the command and send that to the server
            SendPacketToServer(CreatePacket(data));
        }

        /// <summary>
        /// Creates a packet from the given command
        /// </summary>
        /// <param name="input">A command in the form of [28, id, (0-255)]</param>
        /// <returns>Returns a string which can be sent to the HomeCenter server</returns>
        private static string CreatePacket(short[] input)
        {
            // Forms the packet
            int t = input.Length - 1;
            String a = t.ToString();
            for (int n = 0; n < input.Length; n++)
                a += " " + input[n].ToString();
            return a;
        }

        /// <summary>
        /// Sends the packet to the HomeCenter server using the WebSocket initialized earlier
        /// </summary>
        /// <param name="packet">The packet that needs to be sent</param>
        private static void SendPacketToServer(string packet)
        {
            LogHelper.Info("Sending following packet: " + packet);
            ws.Send(packet);
            LogHelper.Info("Packet sent!");
        }
    }
}
