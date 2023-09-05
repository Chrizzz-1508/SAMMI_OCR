using IronOcr;
using IronSoftware.Drawing;
using Newtonsoft.Json;
using System;
using System.Net;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text.RegularExpressions;

namespace SAMMI_OCR
{
    internal class Program
    {
        public static string folderPath = "";
        static void Main(string[] args)
        {
            // Get the path of the currently executing .exe file
            string exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;

            // Get the directory containing the .exe file
            folderPath = System.IO.Path.GetDirectoryName(exePath);

            string scannedFolderPath = Path.Combine(folderPath, "scanned");

            string ip = "127.0.0.1";
            string port = "9450";
            string pw = "";
            string license = "";

            // Read the settings
            if (File.Exists(folderPath + "\\settings.ini"))
            {
                // Read the settings from the INI file using a dictionary.
                Dictionary<string, string> settings = new Dictionary<string, string>();

                foreach (var line in File.ReadLines(folderPath + "\\settings.ini"))
                {
                    if (line.StartsWith("[") && line.EndsWith("]"))
                    {
                        continue;
                    }

                    // Split the line into key and value pairs.
                    string[] parts = line.Split(new char[] { '=' }, StringSplitOptions.RemoveEmptyEntries);

                    if (parts.Length == 2)
                    {
                        // Trim and remove surrounding double quotes (if any).
                        string key = parts[0].Trim();
                        string value = parts[1].Trim('\"');

                        // Add the key-value pair to the settings dictionary.
                        settings[key] = value;
                    }
                }

                // Now you can access the settings.
                if (settings.ContainsKey("ip")) ip = settings["ip"];
                if (settings.ContainsKey("port")) port = settings["port"];
                if (settings.ContainsKey("pw")) pw = settings["pw"];
                if (settings.ContainsKey("license"))
                {
                    license = settings["license"];
                }
                else
                {
                    WriteLog(DateTime.Now.ToString() + " - Error: No license found");
                    return;
                }
            }
            else
            {
                WriteLog(DateTime.Now.ToString() + " - Error: settings.ini not found!");
                return;
            }

            string sOutput = "";

            // Find the newest PNG file
            string[] pngFiles = Directory.GetFiles(folderPath, "*.png");

            if (pngFiles.Length == 0)
            {
                WriteLog(DateTime.Now.ToString() + " - Error: No PNG files found");
                return;
            }

            string newestPngFile = pngFiles
                .Select(file => new FileInfo(file))
                .OrderByDescending(fileInfo => fileInfo.LastWriteTime)
                .First()
                .FullName;

            // Extract values for the crop rectangle from the filename
            string[] fileNameParts = Path.GetFileNameWithoutExtension(newestPngFile).Split('_');
            if (fileNameParts.Length != 5)
            {
                WriteLog(DateTime.Now.ToString() + " - Error: Invalid filename format: " + newestPngFile);
                return;
            }

            int x = int.Parse(fileNameParts[1]);
            int y = int.Parse(fileNameParts[2]);
            int width = int.Parse(fileNameParts[3]);
            int height = int.Parse(fileNameParts[4]);

            // Do Text Recognition
            try
            {
                IronOcr.License.LicenseKey = license;
                var ocr = new IronTesseract();
                using (var ocrInput = new OcrInput())
                {
                    var cropRectangle = new CropRectangle(x, y, width, height);
                    ocrInput.AddImage(newestPngFile, cropRectangle);
                    var ocrResult = ocr.Read(ocrInput);
                    sOutput = ocrResult.Text;
                    WriteLog(DateTime.Now.ToString() + " - Text recognized: " + sOutput);
                }
            }
            catch (Exception ex)
            {
                WriteLog(DateTime.Now.ToString() + " - Error: " + ex.Message);
                return;
            }


            // Create the "scanned" folder if it doesn't exist
            Directory.CreateDirectory(scannedFolderPath);

            // Move the processed image file to the "scanned" folder
            string scannedFilePath = Path.Combine(scannedFolderPath, Path.GetFileName(newestPngFile));
            File.Move(newestPngFile, scannedFilePath);

            // Disable Expect100Continue (needed for SAMMI)
            System.Net.ServicePointManager.Expect100Continue = false;

            // Define the JSON body object
            var data = new
            {
                trigger = "SAMMI OCR",
                value = sOutput
            };

            // Serialize the JSON body object into a string
            var json = JsonConvert.SerializeObject(data);

            // Send data to SAMMI via webhook
            var httpRequest = (HttpWebRequest)WebRequest.Create("http://" + ip + ":" + port + "/webhook");
            httpRequest.Method = "POST";
            httpRequest.ContentType = "application/json";

            if (!String.IsNullOrEmpty(pw))
            {
                httpRequest.Headers["Authorization"] = pw;
            }

            using (var streamWriter = new StreamWriter(httpRequest.GetRequestStream()))
            {
                streamWriter.Write(json);
            }

            try
            {
                var httpResponse = (HttpWebResponse)httpRequest.GetResponse();
            }
            catch (Exception ex)
            {
                WriteLog(DateTime.Now.ToString() + " - Error: " + ex.Message);
            }
        }
        private static void WriteLog(string sMessage)
        {
            using (StreamWriter sw = new StreamWriter(folderPath + "\\" + "ocr.log", true))
            {
                sw.WriteLine(sMessage);
                sw.Flush();
                sw.Close();
            }
        }
    }
}