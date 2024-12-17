using AutoUpdaterDotNET;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace AutoUpdaterTest
{
    public class Serializer
    {
        /// <summary>
        /// populate a class with xml data 
        /// </summary>
        /// <typeparam name="T">Object Type</typeparam>
        /// <param name="input">xml data</param>
        /// <returns>Object Type</returns>
        public T Deserialize<T>(string input) where T : class
        {
            System.Xml.Serialization.XmlSerializer ser = new System.Xml.Serialization.XmlSerializer(typeof(T));

            using (StringReader sr = new StringReader(input))
            {
                return (T)ser.Deserialize(sr);
            }
        }

        /// <summary>
        /// convert object to xml string
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="ObjectToSerialize"></param>
        /// <returns></returns>
        public string Serialize<T>(T ObjectToSerialize)
        {
            XmlSerializer xmlSerializer = new XmlSerializer(ObjectToSerialize.GetType());

            using (StringWriter textWriter = new StringWriter())
            {
                xmlSerializer.Serialize(textWriter, ObjectToSerialize);
                return textWriter.ToString();
            }
        }

    }

    static public class AutoUpdateTools
    {
        static bool IsJson;
        static AutoUpdateTools()
        {
            AutoUpdater.CheckForUpdateEvent += AutoUpdate_CheckForUpdateEvent;
            AutoUpdater.ParseUpdateInfoEvent += AutoUpdater_ParseUpdateInfoEvent;
        }

        private static void AutoUpdater_ParseUpdateInfoEvent(ParseUpdateInfoEventArgs args)
        {
            try
            {
                if (!IsJson)
                {
                    Serializer ser = new Serializer();
                    args.UpdateInfo = ser.Deserialize<UpdateInfoEventArgs>(args.RemoteData);
                }
                else
                {
                    var json = JsonConvert.DeserializeObject<JObject>(args.RemoteData);             
                    args.UpdateInfo = new UpdateInfoEventArgs
                    {
                        CurrentVersion = json.ContainsKey("version") ? json["version"].ToString() : null,
                        ChangelogURL = json.ContainsKey("changelog") ? json["changelog"].ToString() : null,
                        DownloadURL = json.ContainsKey("url") ? json["url"].ToString() : null,
                        Mandatory = json.ContainsKey("mandatory") ? new Mandatory
                        {
                            Value = json["mandatory"]["value"].Value<bool>(),
                            UpdateMode = (Mode)json["mandatory"]["mode"].Value<int>(),
                            MinimumVersion = json["mandatory"]["minVersion"].ToString()
                        } : null,
                        CheckSum = json.ContainsKey("checksum") ? new CheckSum
                        {
                            Value = json["checksum"]["value"].ToString(),
                            HashingAlgorithm = json["checksum"]["hashingAlgorithm"].ToString()
                        } : null
                    };
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, e.GetType().ToString(), MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }
        }

        private static void AutoUpdate_CheckForUpdateEvent(UpdateInfoEventArgs args)
        {
            if (args != null)
            {
                if (args.IsUpdateAvailable && args.CurrentVersion.CompareTo(args.Mandatory.MinimumVersion) >= 0)
                {
                    DialogResult dialogResult;
                    if (args.Mandatory.Value)
                    {
                        dialogResult =
                            MessageBox.Show(
                                $@"There is new version {args.CurrentVersion} available. You are using version {args.InstalledVersion}. \n
This is required update. Press Ok to begin updating the application.", @"Update Available",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                    }
                    else
                    {
                        dialogResult =
                            MessageBox.Show(
                                $@"There is new version {args.CurrentVersion} available. You are using version {args.InstalledVersion}. \n
Do you want to update the application now?", @"Update Available",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Information);
                    }
                    if (dialogResult.Equals(DialogResult.Yes) || dialogResult.Equals(DialogResult.OK))
                    {
                        try
                        {
                            if (AutoUpdater.DownloadUpdate(args))
                            {
                                Application.Exit();
                            }
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show(e.Message, e.GetType().ToString(), MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                        }
                    }
                }
                else
                {
                    MessageBox.Show(@"There is no update available please try again later.", @"No update available",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                if (args.Error is WebException)
                {
                    MessageBox.Show(
                        @"There is a problem reaching update server. Please check your internet connection and try again later.",
                        @"Update Check Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show(args.Error.Message,
                        args.Error.GetType().ToString(), MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        public static void StartUpdaterAsync(string updatePath)
        {
            if (updatePath.Contains(".json"))
                IsJson = true;          
            AutoUpdater.Start(updatePath);
        }
    }
}
