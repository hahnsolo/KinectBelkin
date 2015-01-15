//------------------------------------------------------------------------------
// <copyright file="MainWindow.xaml.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

namespace Microsoft.Samples.Kinect.SpeechBasics
{
    using System;
    using System.Collections.Generic;    
    using System.ComponentModel;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Text;    
    using System.Windows;    
    using System.Windows.Documents;
    using System.Windows.Media;
    using Microsoft.Kinect;    
    using Microsoft.Speech.AudioFormat;
    using Microsoft.Speech.Recognition;
    using System.Net.Mail;
    using System.Diagnostics;

    /// <summary>
    /// Interaction logic for MainWindow
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1001:TypesThatOwnDisposableFieldsShouldBeDisposable",
        Justification = "In a full-fledged application, the SpeechRecognitionEngine object should be properly disposed. For the sake of simplicity, we're omitting that code in this sample.")]
    public partial class MainWindow : Window
    {
        
        /// <summary>
        /// Active Kinect sensor.
        /// </summary>
        private KinectSensor kinectSensor = null;


        /// <summary>
        /// Active auto detects
        /// </summary>
        string stuff = "";
        List<String> names = new List<String>();
        List<int> state = new List<int>();

        /// <summary>
        /// Stream for 32b-16b conversion.
        /// </summary>
        private KinectAudioStream convertStream = null;

        /// <summary>
        /// Speech recognition engine using audio data from Kinect.
        /// </summary>
        private SpeechRecognitionEngine speechEngine = null;

     

 
        List<int> counterList;
        int counter2;

        /// <summary>
        /// Initializes a new instance of the MainWindow class.
        /// </summary>
        public MainWindow()
        {
            counterList = new List<int>();
            counter2 = 0;
            this.InitializeComponent();
        }

        /// <summary>
        /// Enumeration of directions in which turtle may be facing.
        /// </summary>
      
        /// <summary>
        /// Gets the metadata for the speech recognizer (acoustic model) most suitable to
        /// process audio from Kinect device.
        /// </summary>
        /// <returns>
        /// RecognizerInfo if found, <code>null</code> otherwise.
        /// </returns>
        private static RecognizerInfo TryGetKinectRecognizer()
        {
            IEnumerable<RecognizerInfo> recognizers;
            
            // This is required to catch the case when an expected recognizer is not installed.
            // By default - the x86 Speech Runtime is always expected. 
            try
            {
                recognizers = SpeechRecognitionEngine.InstalledRecognizers();
            }
            catch (COMException)
            {
                return null;
            }

            foreach (RecognizerInfo recognizer in recognizers)
            {
                string value;
                recognizer.AdditionalInfo.TryGetValue("Kinect", out value);
                if ("True".Equals(value, StringComparison.OrdinalIgnoreCase) && "en-US".Equals(recognizer.Culture.Name, StringComparison.OrdinalIgnoreCase))
                {
                    return recognizer;
                }
            }

            return null;
        }

        /// <summary>
        /// Execute initialization tasks.
        /// </summary>
        /// <param name="sender">object sending the event</param>
        /// <param name="e">event arguments</param>
        private void WindowLoaded(object sender, RoutedEventArgs e)
        {
            // Only one sensor is supported
            this.kinectSensor = KinectSensor.GetDefault();

            if (this.kinectSensor != null)
            {
                // open the sensor
                this.kinectSensor.Open();

                // grab the audio stream
                IReadOnlyList<AudioBeam> audioBeamList = this.kinectSensor.AudioSource.AudioBeams;
                System.IO.Stream audioStream = audioBeamList[0].OpenInputStream();

                // create the convert stream
                this.convertStream = new KinectAudioStream(audioStream);
            }
            else
            {
                // on failure, set the status text
       
                return;
            }
            System.Diagnostics.Process myProcess2 = new System.Diagnostics.Process();
            ProcessStartInfo myProcessStartInfo2 = new ProcessStartInfo("cmd.exe", @"/c C:\python27\scripts\wemo.exe clear");
            myProcessStartInfo2.UseShellExecute = false;
            myProcessStartInfo2.RedirectStandardOutput = true;
            myProcess2.StartInfo = myProcessStartInfo2;
            myProcess2.Start();

            
            System.Diagnostics.Process myProcess = new System.Diagnostics.Process();
            ProcessStartInfo myProcessStartInfo = new ProcessStartInfo("cmd.exe", @"/c C:\python27\scripts\wemo.exe status");
            myProcessStartInfo.UseShellExecute = false;
            myProcessStartInfo.RedirectStandardOutput = true;
            myProcess.StartInfo = myProcessStartInfo;
                while (myProcess2.HasExited == false)
                {


                }
            myProcess.Start();
            
            StreamReader myStreamReader = myProcess.StandardOutput;
          
            //System.Diagnostics.Process.Start("cmd.exe", @"/c C:\python27\scripts\wemo.exe status");
            while (!myStreamReader.EndOfStream)
            {
                stuff = myStreamReader.ReadLine();
                int stringend = 0;
                int stateend = 0;
                stringend = stuff.Length - 11;
                stateend = stuff.Length - 1;
                names.Add(stuff.Substring(8, stringend));
                state.Add(Convert.ToInt16(stuff.Substring(stateend, 1)));

            }


            RecognizerInfo ri = TryGetKinectRecognizer();

            if (null != ri)
            {
            

                this.speechEngine = new SpeechRecognitionEngine(ri.Id);

                /****************************************************************
                * 
                * Use this code to create grammar programmatically rather than from
                * a grammar file.
                * 
                * var directions = new Choices();
                * directions.Add(new SemanticResultValue("straight", "FORWARD"));
                * directions.Add(new SemanticResultValue("backward", "BACKWARD"));
                * directions.Add(new SemanticResultValue("backwards", "BACKWARD"));
                * directions.Add(new SemanticResultValue("back", "BACKWARD"));
                * directions.Add(new SemanticResultValue("turn left", "LEFT"));
                * directions.Add(new SemanticResultValue("turn right", "RIGHT"));
                *
                * var gb = new GrammarBuilder { Culture = ri.Culture };
                * gb.Append(directions);
                *
                * var g = new Grammar(gb);
                * 
                ****************************************************************/

                // Create a grammar from grammar definition XML file.
               /* using (var memoryStream = new MemoryStream(Encoding.ASCII.GetBytes(Properties.Resources.SpeechGrammar)))
                {
                    var g = new Grammar(memoryStream);
                    this.speechEngine.LoadGrammar(g);
                }*/

                
                var directions = new Choices();
                names.ForEach(delegate(String name)
                {
                    Console.Write("I am Here");
                    Console.WriteLine(name);
                    directions.Add(new SemanticResultValue(name,name));
                });
                directions.Add("No Devices");
                var gb = new GrammarBuilder { Culture = ri.Culture };
                gb.Append(directions);
                var g = new Grammar(gb);
                this.speechEngine.LoadGrammar(g);
                Console.WriteLine(this.speechEngine.Grammars.Count);
                this.speechEngine.SpeechRecognized += this.SpeechRecognized;
                this.speechEngine.SpeechRecognitionRejected += this.SpeechRejected;
        
                // let the convertStream know speech is going active
                this.convertStream.SpeechActive = true;

                // For long recognition sessions (a few hours or more), it may be beneficial to turn off adaptation of the acoustic model. 
                // This will prevent recognition accuracy from degrading over time.
                ////speechEngine.UpdateRecognizerSetting("AdaptationOn", 0);

                this.speechEngine.SetInputToAudioStream(
                this.convertStream, new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
                this.speechEngine.RecognizeAsync(RecognizeMode.Multiple);
            }
            else
            {    
            }
        }
        /// <summary>
        /// Execute un-initialization tasks.
        /// </summary>
        /// <param name="sender">object sending the event.</param>
        /// <param name="e">event arguments.</param>
        private void WindowClosing(object sender, CancelEventArgs e)
        {
            if (null != this.convertStream)
            {
                this.convertStream.SpeechActive = false;
            }
            if (null != this.speechEngine)
            {
                this.speechEngine.SpeechRecognized -= this.SpeechRecognized;
                this.speechEngine.SpeechRecognitionRejected -= this.SpeechRejected;
                this.speechEngine.RecognizeAsyncStop();
            }
            if (null != this.kinectSensor)
            {
                this.kinectSensor.Close();
                this.kinectSensor = null;
            }
        }

        /// <summary>
        /// Remove any highlighting from recognition instructions.
        /// </summary>
        private void ClearRecognitionHighlights()
        {
         
        }

        /// <summary>
        /// Handler for recognized speech events.
        /// </summary>
        /// <param name="sender">object sending the event.</param>
        /// <param name="e">event arguments.</param>
        private void SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            // Speech utterance confidence below which we treat speech as if it hadn't been heard
            const double ConfidenceThreshold = 0.4;

            if (e.Result.Confidence >= ConfidenceThreshold)
            {
                names.ForEach(delegate(String name)
                {
                
                    if (e.Result.Semantics.Value.ToString() == name)
                    {
                        Console.WriteLine("I am in the for each");
                        System.Diagnostics.Process myProcess = new System.Diagnostics.Process();
                        StringBuilder sa = new StringBuilder();
                        sa.Append(@"/c c:\python27\scripts\wemo.exe switch "+name.ToString()+" toggle");
                        Console.Write(sa.ToString());
                        ProcessStartInfo myProcessStartInfo = new ProcessStartInfo("cmd.exe",sa.ToString());
                        myProcessStartInfo.UseShellExecute = false;
                        myProcessStartInfo.RedirectStandardOutput = true;
                        myProcess.StartInfo = myProcessStartInfo;
                        myProcess.Start();
                    }
                });
               
            }
        }
        /// <summary>
        /// Handler for rejected speech events.
        /// </summary>
        /// <param name="sender">object sending the event.</param>
        /// <param name="e">event arguments.</param>
        private void SpeechRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
          
        }
    }
}