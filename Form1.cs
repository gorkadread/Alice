using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;

namespace Alice
{
    using System.Globalization;
    using System.Speech.Recognition;
    using System.Speech.Synthesis;
    using System.Text.RegularExpressions;
    using System.Threading;
    using Microsoft.Win32;

    public partial class Form1 : Form
    {
        // Globals
        private SpeechRecognitionEngine recognitionEngine;
        List<ProcessObject> processesRunning = new List<ProcessObject>();
        List<ApplicationObject> installedApplications = new List<ApplicationObject>();
        SpeechSynthesizer reader = new SpeechSynthesizer();

        public Form1()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-GB");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-GB");
            CultureInfo culture = new CultureInfo("en-GB");
            recognitionEngine = new SpeechRecognitionEngine(culture);
            recognitionEngine.SetInputToDefaultAudioDevice();

            recognitionEngine.SpeechRecognized += (s, args) =>
            {
                string line = string.Empty;
                float security = 1.0f;
                foreach (RecognizedWordUnit word in args.Result.Words)
                {
                    if (word.Confidence > 0.01f)
                    {
                        line += word.Text + " ";
                        security = security * word.Confidence;
                    }
                }

                if (line.Contains("Alice"))
                {
                    var command = Regex.Replace(line, "Alice", string.Empty).Trim();

                    var adress = "www.google.se";

                    var p = new Process();

                    switch (command)
                    {
                        case "mute":
                            if (reader != null)
                            {
                                reader.Dispose();
                            }
                            break;
                        case "run notepad":
                            reader.SpeakAsync("Starting notepad");
                            p = Process.Start("notepad.exe");
                            processesRunning.Add(
                                new ProcessObject { ProcessName = "Notepad", ProcessId = p.Id, Started = DateTime.Now });
                            break;
                        case "run calculator":
                            reader.SpeakAsync("Starting calculator");
                            p = Process.Start("calc.exe");
                            processesRunning.Add(
                                new ProcessObject { ProcessName = "Calculator", ProcessId = p.Id, Started = DateTime.Now });
                            break;
                        case "run paint":
                            reader.SpeakAsync("Starting paint");
                            p = Process.Start("mspaint.exe");
                            processesRunning.Add(
                                new ProcessObject { ProcessName = "Paint", ProcessId = p.Id, Started = DateTime.Now });
                            break;
                        case "run firefox":
                            reader.SpeakAsync("Starting firefox");
                            p = Process.Start(@"C:\Program Files (x86)\Mozilla Firefox\firefox.exe");
                            processesRunning.Add(
                                new ProcessObject { ProcessName = "Firefox", ProcessId = p.Id, Started = DateTime.Now });
                            break;
                        case "run browser":
                            reader.SpeakAsync("Starting default browser");
                            string defaultBrowserPath = GetDefaultBrowserPath();
                            p = Process.Start(defaultBrowserPath);
                            processesRunning.Add(
                                new ProcessObject { ProcessName = "Browser", ProcessId = p.Id, Started = DateTime.Now });
                            break;
                        case "kill notepad":
                            reader.SpeakAsync("Killing notepad");
                            try
                            {
                                ProcessObject matches =
                                    processesRunning.Where(item => item.ProcessName == "Notepad").OrderByDescending(
                                        process => process.Started).First();
                                foreach (Process proc in Process.GetProcesses())
                                {
                                    if (proc.Id == matches.ProcessId)
                                    {
                                        proc.Kill();
                                        processesRunning.RemoveAll(x => x.ProcessId == matches.ProcessId);
                                        reader.SpeakAsync("Done");
                                    }
                                }
                            }
                            catch
                            {
                                reader.SpeakAsync("Could not kill notepad");
                            }
                            break;
                        case "kill all notepad":
                            reader.SpeakAsync("Killing all notepads");
                            try
                            {
                                List<ProcessObject> matches =
                                    this.processesRunning.Where(proc => proc.ProcessName == "Notepad").ToList();

                                foreach (Process proc in Process.GetProcesses())
                                {
                                    foreach (ProcessObject obj in matches)
                                        if (proc.Id == obj.ProcessId)
                                        {
                                            proc.Kill();
                                            processesRunning.RemoveAll(x => x.ProcessId == obj.ProcessId);
                                        }
                                }
                                reader.SpeakAsync("Done");
                            }
                            catch
                            {
                                reader.SpeakAsync("Could not kill all notepads");
                            }
                            break;
                        case "kill calculator":
                            reader.SpeakAsync("Killing calculator");
                            try
                            {
                                ProcessObject matches =
                                    processesRunning.Where(item => item.ProcessName == "Calculator").
                                        OrderByDescending(process => process.Started).First();
                                foreach (Process proc in Process.GetProcesses())
                                {
                                    if (proc.Id == matches.ProcessId)
                                    {
                                        proc.Kill();
                                        processesRunning.RemoveAll(x => x.ProcessId == matches.ProcessId);
                                        reader.SpeakAsync("Done");
                                    }
                                }
                            }
                            catch
                            {
                                reader.SpeakAsync("Could not kill calculator");
                            }
                            break;
                        case "kill all calculator":
                            reader.SpeakAsync("Killing all calculators");
                            try
                            {
                                List<ProcessObject> matches =
                                    this.processesRunning.Where(proc => proc.ProcessName == "Calculator").ToList();
                                foreach (Process proc in Process.GetProcesses())
                                {
                                    foreach (ProcessObject obj in matches)
                                        if (proc.Id == obj.ProcessId)
                                        {
                                            proc.Kill();
                                            processesRunning.RemoveAll(x => x.ProcessId == obj.ProcessId);
                                        }
                                }
                                reader.SpeakAsync("Done");
                            }
                            catch
                            {
                                reader.SpeakAsync("Could not kill all calculators");
                            }
                            break;
                        case "kill paint":
                            reader.SpeakAsync("Killing paint");
                            try
                            {
                                ProcessObject matches =
                                    processesRunning.Where(item => item.ProcessName == "Paint").OrderByDescending(
                                        process => process.Started).First();
                                foreach (Process proc in Process.GetProcesses())
                                {
                                    if (proc.Id == matches.ProcessId)
                                    {
                                        proc.Kill();
                                        processesRunning.RemoveAll(x => x.ProcessId == matches.ProcessId);
                                        reader.SpeakAsync("Done");
                                    }
                                }
                            }
                            catch
                            {
                                reader.SpeakAsync("Could not kill paint");
                            }
                            break;
                        case "kill all paint":
                            reader.SpeakAsync("Killing all paints");
                            try
                            {
                                List<ProcessObject> matches =
                                    this.processesRunning.Where(proc => proc.ProcessName == "Paint").ToList();

                                foreach (Process proc in Process.GetProcesses())
                                {
                                    foreach (ProcessObject obj in matches)
                                        if (proc.Id == obj.ProcessId)
                                        {
                                            proc.Kill();
                                            processesRunning.RemoveAll(x => x.ProcessId == obj.ProcessId);
                                        }
                                }
                                reader.SpeakAsync("Done");
                            }
                            catch
                            {
                                reader.SpeakAsync("Could not kill all paints");
                            }
                            break;
                        // Firefox only uses one processid
                        case "kill firefox":
                        case "kill all firefox":
                            reader.SpeakAsync("Killing firefox");
                            try
                            {
                                ProcessObject matches =
                                    processesRunning.Where(item => item.ProcessName == "Firefox").First();
                                foreach (Process proc in Process.GetProcesses())
                                {
                                    if (proc.Id == matches.ProcessId)
                                    {
                                        proc.Kill();
                                        processesRunning.RemoveAll(x => x.ProcessName == matches.ProcessName);
                                        reader.SpeakAsync("Done");
                                    }
                                }
                            }
                            catch
                            {
                                reader.SpeakAsync("Could not kill firefox");
                            }
                            break;
                        case "kill browser":
                            reader.SpeakAsync("Killing default browser");
                            try
                            {
                                ProcessObject matches =
                                    processesRunning.Where(item => item.ProcessName == "Browser").First();
                                foreach (Process proc in Process.GetProcesses())
                                {
                                    if (proc.Id == matches.ProcessId)
                                    {
                                        proc.Kill();
                                        processesRunning.RemoveAll(x => x.ProcessName == matches.ProcessName);
                                        reader.SpeakAsync("Done");
                                    }
                                }
                            }
                            catch
                            {
                                reader.SpeakAsync("Could not kill default browser");
                            }
                            break;
                    }

                    richTextBox3.Text += line + "certainty: " + security * 100 + "%";
                    richTextBox3.Text += Environment.NewLine;
                    var listOfProcesses = "";
                    foreach (ProcessObject pr in processesRunning)
                    {
                        listOfProcesses += pr.ProcessName + " " + pr.ProcessId + Environment.NewLine;
                    }
                    richTextBox2.Text = "Programs in list: " + Environment.NewLine + listOfProcesses;
                }
            };
            recognitionEngine.LoadGrammar(CreateGrammarObject());

            InitializeComponent();

            // Get all installed applications
            GetInstalledApplications();

            recognitionEngine.RecognizeAsync(RecognizeMode.Multiple);
        }

        private void GetInstalledApplications()
        {
            string registry_key = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            using (Microsoft.Win32.RegistryKey key = Registry.LocalMachine.OpenSubKey(registry_key))
            {
                var listOfApplications = "";
                foreach (string subkey_name in key.GetSubKeyNames())
                {
                    using (RegistryKey subkey = key.OpenSubKey(subkey_name))
                    {
                        if ((string)subkey.GetValue("DisplayName") != null && (string)subkey.GetValue("InstallLocation") != null)
                        {
                            installedApplications.Add(
                                new ApplicationObject { ApplicationName = (string)subkey.GetValue("DisplayName"), ApplicationPath = (string)subkey.GetValue("InstallLocation") });

                            listOfApplications += (string)subkey.GetValue("DisplayName") + " (" + (string)subkey.GetValue("InstallLocation") + ")" + Environment.NewLine;
                        }
                    }
                }

                richTextBox1.Text = listOfApplications;
            }
        }

        private string GetDefaultBrowserPath()
        {
            RegistryKey regkey = null;
            string browser = string.Empty;

            try
            {
                //set the registry key we want to open
                regkey = Registry.ClassesRoot.OpenSubKey(@"HTTP\shell\open\command", false);

                //get rid of the enclosing quotes
                browser = regkey.GetValue(null).ToString().ToLower().Replace("" + (char)34, "");
                //check to see if the value ends with .exe (this way we can remove any command line arguments)
                if (!browser.EndsWith("exe"))
                    //get rid of all command line arguments (anything after the .exe must go)
                    browser = browser.Substring(0, browser.LastIndexOf(".exe") + 4);
                var test = browser;
            }
            catch (Exception ex)
            {
                browser = string.Format("ERROR: An exception of type: {0} occurred in method: {1} in the following module: {2}", ex.GetType(), ex.TargetSite, this.GetType());
            }
            finally
            {
                //check and see if the key is still open, if so
                //then close it
                if (regkey != null)
                    regkey.Close();
            }
            //return the value
            return browser;
        }

        private Grammar CreateGrammarObject()
        {
            Choices commandChoices = new Choices("run calculator", "run notepad", "run firefox", "run paint", "run browser", "kill calculator", "kill all calculator", "kill notepad", "kill all notepad", "kill firefox", "kill all firefox", "kill paint", "kill all paint", "kill browser", "mute");
            GrammarBuilder grammarBuilder = new GrammarBuilder("Alice");
            grammarBuilder.Append(commandChoices);
            Grammar g = new Grammar(grammarBuilder);
            return g;
        }
    }
}