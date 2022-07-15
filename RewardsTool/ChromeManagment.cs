using System;
using System.Linq;
using System.IO.Compression;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO;
using System.Windows.Forms;
using System.Security;
using System.Net.Http;
using System.Threading.Tasks;

namespace RewardsEdge
{
    enum OSList
    {
        Windows64,
        Windows32,
        WindowsARM ,
    }

    /**
     * <summary> Exception to raise in case the program can't find the selected Chrome progile. </summary>
     */
    class ProfileNotFound : Exception
    {
        public ProfileNotFound(string message) : base(message) { }
    }


    /**
     * <summary> Exception to raise in case the os is not supported. </summary>
     */
    class InvalidPlatform : Exception
    {
        public InvalidPlatform(string message) : base(message) { }
    }


    class ChromeManagment
    {

        public static OSList currentOS { get; private set; }


        /**
         * <summary>Check if the passed folder exists and consequentially if chrome folder exists too</summary>
         * <param name="path"> The chrome data path, usually in %localappdata%\Google\Chrome\User Data</param>
         * <param name="profileFolder"> The folder to check if it's present</param>
         */
        private static void ExistChromeFolder(string path, string profileFolder)
        {
            if (!Directory.Exists(path))
            {
                Console.WriteLine("The selected profile doesn't exist, insert a valid profile or leave it empty");
                Console.ReadKey();
                throw new ProfileNotFound("Profile " + profileFolder + " not found");
            }
        }


        /**
         * <summary> Check if the user exists, if it doesn't exist it will execute <see cref="ExistChromeFolder(string, string)">ExistChromeFolder</see> function </summary>
         * <param name="profileFolder"> </param>
         */

        private static string ResolveChromeFolder(string profileFolder) {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\Google\Chrome\User Data";
            ExistChromeFolder(path, profileFolder);
            return path;

        }


        /** <summary>Elaborates the given arguments. </summary>
         * <param name="args"> The array of arguments, it must have the same structure of the args passed in <see cref="Main(string[])">Main</see>. </param>
         */
        public static Tuple<string, string, string> Arguments(string[] args)
        {
            
            string driverPath = @".\";
            string profileFolder = "Profile 1";
            string chromePath = ResolveChromeFolder(profileFolder);
            bool _w = true;
            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "-w":
                        if (_w)
                        {
                            Console.WriteLine("Press a button to start the program");
                            Console.ReadKey();
                            _w = false;
                        }
                        break;
                    case "-u":
                        if (args.Length - 1 > i && args[i + 1][0] != '-')
                        {
                            profileFolder = args[++i];
                            chromePath = ResolveChromeFolder(profileFolder);

                        }
                        break;

                    case "-p":
                        if (args.Length - 1 > i && args[i + 1][0] != '-')
                        {
                            driverPath = args[++i];
                        }
                        break;
                }

            }
            if (driverPath.Last() != '\\')
                driverPath += "\\";
            return Tuple.Create(profileFolder, chromePath, driverPath);
        }

        private static async Task<string> GetDriverVersion(string version) {
            HttpClient httpClient = new HttpClient();
            string driverVersion="";
            bool loop;
            do {
                loop = true;
                int pPos = version.LastIndexOf('.');
                if (pPos >= 0) {
                    loop = false;
                    version = version.Substring(0, pPos);
                }
                var resp = await httpClient.GetAsync("https://chromedriver.storage.googleapis.com/LATEST_RELEASE" + (pPos >= 0 ? "_" + version : ""));
                if(resp.IsSuccessStatusCode) {
                    driverVersion = await resp.Content.ReadAsStringAsync();
                    loop = false;
                }
            } while (loop);
            return driverVersion;
        }

        /**
         * <summary> Download the right driver version for Chrome.</summary>
         * To get the current Chrome version it is used the function <see cref="GetChromeVersion">GetChromeVersion</see>.
         * If in the folder there is a "chromedriver_win32.zip" file the program will ends, it is necessary to remove that file.
         * If there is already a "chromedriver.exe" file and the program can't remove it the program will ends, it is necessary to remove that file.
         * <param name="path"> The path where download the driver</param>
         */
        public static void DownloadDriver(string path)
        {
            string actualVersion = GetChromeVersion();
            string version;
            var task = GetDriverVersion(actualVersion);
            version = task.Result;
            string req = "https://chromedriver.storage.googleapis.com/" + version + "/chromedriver_" + GetOSArch() + ".zip";
            string zipPath = Path.GetFullPath(path + "chromedriver_win32.zip");
            string exePath = Path.GetFullPath(path + "chromedriver.exe");


            if (File.Exists(zipPath))
            {
                Console.WriteLine(zipPath + " already exists, please remove it and run again the program");
                Application.Exit();
            }

            if (File.Exists(exePath))
            {
                Console.WriteLine("Removing old driver version" + exePath);
                try
                {
                    File.Delete(exePath);
                }
                catch (UnauthorizedAccessException)
                {
                    Console.WriteLine("Impossible to remove the driver, please remove it manually and retry");
                    return;
                }

            }

            Console.WriteLine("Downloading zip in " + zipPath);
            using (var client = new System.Net.WebClient())
            {
                client.DownloadFile(req, zipPath);
            }


            using (ZipArchive archive = ZipFile.OpenRead(zipPath))
            {
                Console.WriteLine("Unzipping new driver");
                foreach (ZipArchiveEntry entry in archive.Entries.Where(e => e.FullName == "chromedriver.exe"))
                {
                    entry.ExtractToFile(exePath);
                }
            }
            Console.WriteLine("Removing the zip");
            File.Delete(zipPath);
            Console.WriteLine("Download completed");
        }


        /**
         * <summary> Get the os and architecture used, necessary to download the correct driver. </summary>
         * <returns> The driver platform.</returns>
         */
        private static string GetOSArch()
        {
            string toRet;
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                //string arch = Environment.GetEnvironmentVariable("PROCESSOR_ARCHITEW6432");
                string arch = Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE");
                switch (arch)
                {
                    case "x86":
                        currentOS = OSList.Windows32;
                        toRet = "win32";
                        break;

                    case "AMD64":
                        currentOS = OSList.Windows64;
                        //there is no win64 version
                        toRet = "win32";
                        break;

                    case "ARM64":
                        currentOS = OSList.WindowsARM;
                        //there is no arm64 version
                        toRet = "win32";
                        break;

                    default:
                        throw new InvalidPlatform("This Windows version is not supported");

                }
            }
            else
            {
                throw new InvalidPlatform("Platform not recognized");
            }

            return toRet;
        }


        /**
         * <summary> Gets the Chrome version.</summary>
         * It uses powershell.exe to get the chrome version.
         * <returns> Chrome version.</returns>
         */
        private static string GetChromeVersion() {
            // get version using powershell
            Object ret;
            try {
                 ret= Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon", "version", null);
            }
            catch (SecurityException) {
                Console.WriteLine("Impossible get the chrome version");
                return "0.0.0.0";
            }
            if(ret is null) {
                Console.WriteLine("Impossible get the chrome version");
                return "0.0.0.0";
            }
            else {
                Console.WriteLine("Chrome version: "+ ret.ToString());
                return ret.ToString();
            }
            // get return value
        }
    }
}
