using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;
using nsj = Newtonsoft.Json;
using System.IO.Compression;
using System.Threading;
using Newtonsoft.Json;

namespace ChromeDriverDownloader
{
    public static class ChromeDriverDownloader
    {
        #region Classes
        public class Chrome
        {
            public string platform { get; set; }
            public string url { get; set; }
        }

        public class Chromedriver
        {
            public string platform { get; set; }
            public string url { get; set; }
        }

        public class Downloads
        {
            public List<Chrome> chrome { get; set; }
            public List<Chromedriver> chromedriver { get; set; }
        }

        public class Root
        {
            public DateTime timestamp { get; set; }
            public List<Version> versions { get; set; }
        }

        public class Version
        {
            public string version { get; set; }
            public string revision { get; set; }
            public Downloads downloads { get; set; }
        }
        #endregion


        private static string VersionTextFile = "Version.info";
        private static string ChromeDriverPath = AppDomain.CurrentDomain.BaseDirectory;

        public static void BetaDownload(string ChromeDriverFolder = null)
        {
            if (ChromeDriverFolder != null)
            {
                ChromeDriverPath = ChromeDriverFolder;
            }
        }

        public static void Download(string ChromeDriverFolder = null)
        {
            if (ChromeDriverFolder != null)
            {
                ChromeDriverPath = ChromeDriverFolder;
            }

            Console.WriteLine("Checking Chrome Driver");

            var Result = NeedsUpdate();

            if (Result)
            {
                Console.WriteLine("Updating Chrome Driver");

                bool IsSuccessful = Update();

                if (IsSuccessful)
                {
                    Console.WriteLine("Chrome Driver Updated Successfully");
                }
            }
        }

        private static bool NeedsUpdate()
        {
            if (File.Exists(ChromeDriverPath + VersionTextFile) == false)
            {
                Console.WriteLine("No Previous Download Detected");
                return true;
            }

            string Version = File.ReadAllText(ChromeDriverPath + VersionTextFile);

            if (Version != GetShortenedVersionNumber())
            {
                return true;
            }

            else
            {
                return false;
            }
        }

        private static bool BetaUpdate()
        {
            string TargetVersion = GetShortenedVersionNumber();

            string fullVersionNumber = ChromeDriverDownloader.GetFullVersionNumber();
            string semiFullVersionNumber = ChromeDriverDownloader.GetSemiFullVersionNumber();
            string mediumVersionNumber = ChromeDriverDownloader.GetMediumVersionNumber();
            string shortVersionNumber = ChromeDriverDownloader.GetShortenedVersionNumber();

            string testingLink = "https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json";

            string downloadJson = "downloadOptions.json";

            WebClient webClient = new WebClient();
            webClient.DownloadFile(testingLink, downloadJson);
            var deserializedObject = JsonConvert.DeserializeObject<ChromeDriverDownloader.Root>(File.ReadAllText(downloadJson));
            File.Delete(downloadJson);

            string DownloadLink = null;

            string[] trials = new string[] { fullVersionNumber, semiFullVersionNumber, mediumVersionNumber, shortVersionNumber };

            bool isFound = false;

            foreach (var trial in trials)
            {
                Console.WriteLine("Searching Version: " + trial);

                foreach (var download in deserializedObject.versions)
                {
                    if (download.downloads.chromedriver != null)
                    {
                        if (download.version.Contains(trial))
                        {
                            foreach (var chromedriverItem in download.downloads.chromedriver)
                            {
                                if (chromedriverItem.platform.Contains("win32"))
                                {
                                    DownloadLink = chromedriverItem.url;
                                    isFound = true;
                                    break;
                                }
                            }
                        }

                        if (isFound)
                        {
                            break;
                        }
                    }
                }


                if (isFound)
                {
                    Console.WriteLine("Found: " + trial);
                    break;
                }
            }

            if (isFound)
            {
                DownloadFile(DownloadLink, ChromeDriverPath);
                Thread.Sleep(1024 * 3);

                ExtractZipWithOverwrite(ChromeDriverPath + "chromedriver_win32.zip", ChromeDriverPath);

                string[] ZipFiles = GetFilesInZip(ChromeDriverPath + "chromedriver_win32.zip");

                foreach (var zipFile in ZipFiles)
                {
                    if (zipFile.ToLower().Contains("chromedriver") && zipFile.ToLower().EndsWith(".exe"))
                    {
                        File.Move(ChromeDriverPath + zipFile, ChromeDriverPath + Path.GetFileName(zipFile));
                    }
                }

                File.WriteAllText(ChromeDriverPath + VersionTextFile, TargetVersion);

                return true;
            }

            else
            {
                return false;
            }


        }

        private static bool Update()
        {
            try
            {
                #region Cleanup
                // Close If Running
                CloseProcessByExePath(ChromeDriverPath + "chromedriver.exe");

                // Try Delete Previous Zip
                if (File.Exists("chromedriver_win32.zip"))
                {
                    File.Delete("chromedriver_win32.zip");
                }

                // Try Delete Previous Exe
                if (File.Exists("chromedriver.exe"))
                {
                    File.Delete("chromedriver.exe");
                }
                #endregion

                // Get Target Version & Download Options
                string TargetVersion = GetShortenedVersionNumber();
                List<string> DownloadOptions = GetDownloadList();

                string AppropriateVersion = GetAppropriateVersion(TargetVersion, DownloadOptions.ToArray());

                if (AppropriateVersion == null || AppropriateVersion.Length < 3)
                {
                    throw new Exception("No Matching Version Provided by Chromium Website", new Exception("Managed"));
                }

                // Get Correct Download Link -> Then Download To ChromeDriverPath
                string DownloadLink = "https://chromedriver.storage.googleapis.com/" + GetAppropriateVersion(TargetVersion, DownloadOptions.ToArray());
                DownloadFile(DownloadLink, ChromeDriverPath);

                Thread.Sleep(1024 * 3);

                // Extract
                ExtractZipWithOverwrite(ChromeDriverPath + "chromedriver_win32.zip", ChromeDriverPath);

                File.WriteAllText(ChromeDriverPath + VersionTextFile, TargetVersion);

                return true;
            }

            catch (Exception Ex)
            {
                if (Ex.InnerException != null && Ex.InnerException.Message == "Managed")
                {
                    Console.WriteLine("Chrome Driver Update Error: " + Ex.Message);
                    Console.WriteLine("Doing Beta Update");
                    bool BetaSuccess = BetaUpdate();

                    if (BetaSuccess)
                    {
                        Console.WriteLine("Beta Update Success");
                    }

                    else
                    {
                        Console.WriteLine("Beta Update Failed");
                    }
                }

                else
                {
                    Console.WriteLine("Chrome Driver Update Error: " + Ex.Message + Environment.NewLine + Environment.NewLine + Environment.NewLine + Ex.ToString());
                }

                return false;
            }
        }

        public static void ExtractZipWithOverwrite(string zipFilePath, string extractPath)
        {
            using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    string entryPath = Path.Combine(extractPath, entry.FullName);

                    // Check if the file already exists and overwrite it if necessary
                    if (File.Exists(entryPath))
                    {
                        File.Delete(entryPath);
                    }

                    // Create directory if it doesn't exist
                    Directory.CreateDirectory(Path.GetDirectoryName(entryPath));

                    entry.ExtractToFile(entryPath);
                }
            }
        }

        public static string[] GetFilesInZip(string zipFilePath)
        {
            List<string> filePaths = new List<string>();

            using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (!string.IsNullOrEmpty(entry.FullName))
                    {
                        filePaths.Add(entry.FullName);
                    }
                }
            }

            return filePaths.ToArray();
        }

        #region Inner Workings
        public class ChromeDownloads
        {
            public string Key { get; set; }
        }

        private static void CloseProcessByExePath(string exeFilePath)
        {
            Process[] processes = Process.GetProcesses();

            foreach (Process process in processes)
            {
                try
                {
                    string processPath = process.MainModule.FileName;
                    if (string.Equals(processPath, exeFilePath, StringComparison.OrdinalIgnoreCase))
                    {
                        process.Kill();
                        process.WaitForExit();
                        Console.WriteLine($"Process for {exeFilePath} closed successfully.");
                        return;
                    }
                }

                catch
                {

                }
            }

            // Console.WriteLine($"No process for {exeFilePath} found.");
        }

        private static void DownloadFile(string Link, string Path)
        {
            using (var client = new WebClient())
            {
                Console.WriteLine(Link + ": " + Path + "chromedriver_win32.zip");
                client.DownloadFile(Link, Path + "chromedriver_win32.zip");
            }
        }

        private static string GetAppropriateVersion(string TargetVersion, string[] Downloads)
        {
            foreach (var Download in Downloads)
            {
                if (Download.Split('.')[0] == TargetVersion)
                {
                    return Download;
                }
            }

            return null;
        }

        private static string DownloadString(string address)
        {
            string text;
            using (var client = new WebClient())
            {
                text = client.DownloadString(address);
            }
            return text;
        }

        public static string GetFullVersionNumber()
        {
            object path;
            path = Registry.GetValue(@"HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe", "", null);

            if (path != null)
            {
                return FileVersionInfo.GetVersionInfo(path.ToString()).FileVersion;
            }

            return null;
        }

        public static string GetShortenedVersionNumber()
        {
            return GetFullVersionNumber().Split('.')[0];
        }

        public static string GetSemiFullVersionNumber()
        {
            string returnVersion = GetFullVersionNumber();

            var splitNumbers = returnVersion.Split('.');

            if (splitNumbers.Count() == 4)
            {
                returnVersion = splitNumbers[0] + "." + splitNumbers[1] + "." + splitNumbers[2] + "." + splitNumbers[3].Substring(0, 2);
            }

            else
            {
                return GetFullVersionNumber();
            }

            return returnVersion;
        }

        public static string GetMediumVersionNumber()
        {
            string returnVersion = GetFullVersionNumber();

            var splitNumbers = returnVersion.Split('.');

            if (splitNumbers.Count() == 4)
            {
                returnVersion = splitNumbers[0] + "." + splitNumbers[1] + "." + splitNumbers[2];
            }

            else
            {
                return GetFullVersionNumber();
            }

            return returnVersion;
        }

        private static List<string> GetDownloadList(bool Verbose = false)
        {
            List<string> Links = new List<string>();

            #region Parse XML, Parse Json -> filteredString
            var XMLString = DownloadString("https://chromedriver.storage.googleapis.com/");
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(XMLString);
            string jsonString = nsj.JsonConvert.SerializeXmlNode(doc);
            string filteredString = jsonString.Replace("\"?xml\":{\"@version\":\"1.0\",\"@encoding\":\"UTF-8\"},", "").Replace("{\"ListBucketResult\":{\"@xmlns\":\"http://doc.s3.amazonaws.com/2006-03-01\",\"Name\":\"chromedriver\",\"Prefix\":\"\",\"Marker\":\"\",\"IsTruncated\":\"false\",\"Contents\":", "");
            filteredString = filteredString.Substring(0, filteredString.Length - 2);
            #endregion

            ChromeDownloads[] chromeDownloads = nsj.JsonConvert.DeserializeObject<ChromeDownloads[]>(filteredString);

            foreach (var chromeDownload in chromeDownloads)
            {
                if (chromeDownload.Key.EndsWith("win32.zip"))
                {
                    if (Verbose) Console.WriteLine(chromeDownload.Key);
                    Links.Add(chromeDownload.Key);
                }
            }

            return Links;
        }
        #endregion
    }
}
