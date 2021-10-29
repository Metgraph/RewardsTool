using Microsoft.Edge.SeleniumTools;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
//using System.Windows;
using System.Windows.Forms;
using System.IO.Compression;

namespace RewardsEdge
{
    class ProfileNotFound : Exception
    {
        public ProfileNotFound(string message) : base(message) { }
    }
    class Rewards
    {
        private static WebDriverWait wait;
        private static IWebDriver driver;
        private static EdgeOptions options;

        private static Tuple<string, string> Arguments(string[] args)
        {
            string edgeUser = "Default";
            string path = @".\";
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
                    case "-p":
                        if (args.Length - 1 > i && args[i + 1][0] != '-')
                        {
                            edgeUser = args[++i];
                            if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Microsoft\\Edge\\User Data" + edgeUser))
                            {
                                Console.WriteLine("The selected profile doesn't exist, insert a valid profile or leave it empty");
                                Console.ReadKey();
                                throw new ProfileNotFound("Profile " + edgeUser + " not found");
                            }
                        }
                        break;

                    case "-f":
                        if (args.Length - 1 > i && args[i + 1][0] != '-')
                        {
                            path = args[++i];
                        }
                        break;
                }

            }
            if (path.Last() != '\\')
                path += "\\";
            return Tuple.Create(edgeUser, path);
        }

        //Take the daily cards
        private static void DailyCards(int sleep = 1000)
        {
            //get the 3 daily cards
            var listDaily = driver.FindElements(By.XPath("//div[@id='daily-sets']//div[@class='c-card-content']"));

            for (int i = 0; i < 3; i++)
            {
                if (!IsCardDone(listDaily[i]))
                {
                    //click on the card
                    Click(listDaily[i].FindElement(By.XPath(".//div[@class='actionLink x-hidden-vp1']/span")));
                    driver.SwitchTo().Window(driver.WindowHandles.Last());
                    //resolve quiz if exist
                    ResolveQuiz();
                    driver.SwitchTo().Window(driver.WindowHandles[0]);

                }


            }
        }

        //take the not daily cards
        private static void OtherCards(int sleep = 1000)
        {
            //get the other cards
            var listOthers = driver.FindElements(By.XPath("//mee-card-group[@id='more-activities']//div[@class='c-card-content']"));
            foreach (var other in listOthers)
            {
                if (!IsCardDone(other))
                {
                    //click on the card
                    Click(other.FindElement(By.XPath(".//div[@class='actionLink x-hidden-vp1']")));
                    driver.SwitchTo().Window(driver.WindowHandles.Last());
                    //resolve quiz if exist
                    ResolveQuiz();
                    driver.SwitchTo().Window(driver.WindowHandles[0]);
                }


            }
        }

        private static void resolvePunchCard(IWebElement punchcard, int sleep)
        {
            Console.WriteLine("Entrato in punch card");
            //open the punch card page
            punchcard.FindElement(By.XPath(".//a")).Click();
            driver.SwitchTo().Window(driver.WindowHandles.Last());
            //for each quiz in the card
            foreach (var toClick in driver.FindElements(By.XPath("//button[@class='btn-primary btn win-color-border-0 card-button-height pull-left margin-right-24 padding-left-24 padding-right-24']/preceding::a[1]")))
            {
                //if the button will not redirect to a bing page stop the program
                string stringUrl = toClick.GetAttribute("href");
                if (!stringUrl.Contains("https%3A%2F%2Faka.ms") && !stringUrl.Contains("https%3A%2F%2Fwww.bing.com"))
                {
                    Console.WriteLine("exit punchcard");
                    break;
                }

                Console.WriteLine("in punchcard quiz");
                Thread.Sleep(sleep);
                //open the url in another page
                Click(toClick);
                driver.SwitchTo().Window(driver.WindowHandles.Last());
                ResolveQuiz();
                driver.SwitchTo().Window(driver.WindowHandles.Last());
            }
            //close page and return to the rewards page
            driver.Close();
            driver.SwitchTo().Window(driver.WindowHandles[0]);
        }
        private static void PunchCard(int sleep = 1000)
        {
            IWebElement punchCard;
            //try to take the punch card
            try
            {
                punchCard = driver.FindElement(By.XPath("//div[@class='c-carousel f-auto-play f-multi-slide f-scrollable-next f-scrollable-previous']"));

            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("Punch card not found");
                return;
            }

            int i_section = 0;
            var sections = punchCard.FindElements(By.XPath("//section"));
            //for each button there is a punch card
            foreach (var button in punchCard.FindElements(By.XPath("//mee-carousel/div/div[1]/div/button")))
            {
                Click(button);
                IWebElement section = sections[i_section++];

                Console.WriteLine("Nome: " + section.FindElement(By.XPath(".//p[@class='c-subheading ng-binding']")).Text);
                bool completed = true;
                foreach (var check in section.FindElements(By.XPath(".//div[@class='icon-container ng-scope']/span/span")))
                {
                    if (check.GetAttribute("class") != "mee-icon mee-icon-StatusCircleOuter checkmark ng-scope")
                    {
                        completed = false;
                        break;
                    }
                }
                //if punch card isn't completed
                if (!completed)
                {
                    resolvePunchCard(section, sleep);
                }

            }

        }

        //check if the card is completed
        private static bool IsCardDone(IWebElement we)
        {
            return we.FindElements(By.XPath(".//span[@class='mee-icon mee-icon-SkypeCircleCheck']")).Count > 0;
        }

        //when is possible click the element in the page
        private static void Click(IWebElement we)
        {
            //IWebElement element = wait.Until(ExpectedConditions.ElementToBeClickable(we));
            wait.Until(e => we.Displayed && we.Enabled ? we : null);
            we.Click();
        }

        //TOCOMPLETE
        //function to recognize if there is a quiz, a pool or nothing
        private static void ResolveQuiz(int sleep = 3500)
        {
            //wait before analyze page, so page has the time to load and in case it isn't a quiz the time to get the card as completed
            Thread.Sleep(sleep);
            // If there is no quiz exit from function
            var overlay = driver.FindElements(By.XPath("//div[@class='btOverlay']"));
            if (overlay.Count == 0)
            {
                driver.Close();
                return;
            }

            //if there is a pool
            var findPoll = overlay[0].FindElements(By.XPath(".//div[@class='bt_poll']"));
            if (findPoll.Count > 0)
            {
                Click(findPoll[0].FindElement(By.XPath(".//div[@id='btoption0']")));
                Thread.Sleep(500);
                driver.Close();
                return;
            }

            //if there is a quiz
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            if ((bool)js.ExecuteScript("return _w.hasOwnProperty('rewardsQuizRenderInfo')"))
            {
                var findQuiz = overlay[0].FindElements(By.XPath(".//div[@id='quizWelcomeContainer']"));
                if (findQuiz.Count > 0)
                    Click(driver.FindElement(By.XPath(".//input[@id='rqStartQuiz']")));
                DoQuiz();
                driver.Close();
            }

        }

        //recognize which type of quiz it is
        private static void DoQuiz(int sleep = 3500)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            if ((bool)js.ExecuteScript("return _w.rewardsQuizRenderInfo.isListicleQuizType"))
            {
                MultipleAnswerQuiz(sleep);
            }
            else if ((bool)js.ExecuteScript("return _w.rewardsQuizRenderInfo.isWOTQuizType"))
            {
                ThisOrThat(sleep);
            }
            else
            {
                SingleAnswerQuiz(sleep);
            }

        }

        //resolve multiple answer
        private static void MultipleAnswerQuiz(int sleep)
        {
            var maxAndCurrent = getMaxAndCurrent();
            //3 sub-quiz
            Console.WriteLine("current: " + maxAndCurrent[1] + ", max: " + maxAndCurrent[0]);
            for (long i = maxAndCurrent[1]; i <= maxAndCurrent[0]; i++)
            {
                //8 possible answers
                for (int j = 0; j < 8; j++)
                {
                    var slide = driver.FindElement(By.XPath("//div[@class='btOverlay']//div[@class='slide']/div[@id]"));
                    Click(slide);
                    Thread.Sleep(100);
                    //when all the correct answers have been selected this div tag with this class will appear
                    var temp = driver.FindElements(By.XPath("//div[@class='btOverlay']//div[@class='b_promtxt rqQPanel b_hide']"));
                    if (temp.Count > 0)
                        break;
                    //if this is the last sub-quiz there is another way to recognize if quiz is ended
                    else if (i == maxAndCurrent[0])
                    {
                        var temp2 = driver.FindElements(By.XPath("//div[@class='btOverlay']//div[@class='headerMessage']"));
                        if (temp2.Count > 0)
                            break;
                    }
                }
                Thread.Sleep(sleep);
            }
        }

        //resolve single answer quiz
        private static void SingleAnswerQuiz(int sleep)
        {
            var maxAndCurrent = getMaxAndCurrent();
            //3 sub-quiz
            for (long i = maxAndCurrent[1]; i <= maxAndCurrent[0]; i++)
            {
                //4 possible answers
                for (int j = 0; j < 4; j++)
                {
                    var slide = driver.FindElement(By.XPath("//div[@class='btOverlay']//input[@class='rqOption']"));
                    Click(slide);
                    Thread.Sleep(100);
                    //check if i selected the correct answer
                    if (driver.FindElements(By.XPath("//div[@class='btOverlay']//input[@class='rqOption correctAnswer']")).Count > 0)
                        break;
                }
                Thread.Sleep(sleep);
            }

        }

        //resolve this or that quiz
        private static void ThisOrThat(int sleep)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            var maxAndCurrent = getMaxAndCurrent(js);
            //10 questions
            for (long i = maxAndCurrent[1]; i <= maxAndCurrent[0]; i++)
            {
                //this value will be used to calculate which is the correct answer
                string IG = driver.FindElement(By.XPath("//span[@id='nc_iid']")).GetAttribute("_ig");
                //get the correct answer value
                string CorrectAnswer = (string)js.ExecuteScript("return _w.rewardsQuizRenderInfo.correctAnswer");

                var Options = driver.FindElements(By.XPath("//div[@class='btOverlay']//div[@class='btOptionCard' and @id]"));
                //calculate value of first answer
                string firstOptValue = ResolveCorrectAnswer(Options[0].GetAttribute("data-option"), IG);

                if (CorrectAnswer == firstOptValue)
                    //Options[0].Click();
                    Click(Options[0]);
                else
                    //Options[1].Click();
                    Click(Options[1]);
            }
        }

        private static long[] getMaxAndCurrent()
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            return getMaxAndCurrent(js);
        }
        private static long[] getMaxAndCurrent(IJavaScriptExecutor js)
        {
            long[] res = new long[2];
            res[0] = (long)js.ExecuteScript("return _w.rewardsQuizRenderInfo.maxQuestions");
            res[1] = (long)js.ExecuteScript("return _w.rewardsQuizRenderInfo.currentQuestionNumber");
            return res;
        }

        //trasposition in c# of function "br" present in the html code in ThisOrThat quiz
        private static string ResolveCorrectAnswer(string dataOption, string IG)
        {
            int t = 0;
            foreach (char c in dataOption)
            {
                t += c;
            }

            string hex = IG.Substring(IG.Length - 2);
            return (t + Convert.ToInt32(hex, 16)).ToString();
        }

        //make the desktop and mobile searches to get the point
        private static void BingResearches(long pointForSearch, bool closeBrowserAtEnd = false, int length = 4, int sleep = 1000)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            //get how many searches remain for desktop
            long maxPointPC = (long)js.ExecuteScript("return dashboard.userStatus.counters.pcSearch[0].pointProgressMax");
            maxPointPC += (long)js.ExecuteScript("return dashboard.userStatus.counters.pcSearch[1].pointProgressMax");
            long pointPC = (long)js.ExecuteScript("return dashboard.userStatus.counters.pcSearch[0].pointProgress");
            try
            {
                //get how many searches remain for mobile
                long maxPointMobile = (long)js.ExecuteScript("return dashboard.userStatus.counters.mobileSearch[0].pointProgressMax");
                pointPC += (long)js.ExecuteScript("return dashboard.userStatus.counters.pcSearch[1].pointProgress");
                long pointMobile = (long)js.ExecuteScript("return dashboard.userStatus.counters.mobileSearch[0].pointProgress");
            }
            catch (WebDriverException)
            {
            }

            RandomStringGen rsg = new RandomStringGen();
            // Desktop Edge searches
            for (long i = 0; i < maxPointPC - pointPC; i += pointForSearch)
            {
                driver.Navigate().GoToUrl("https://bing.com/search?q=" + rsg.GenString(length));
                Thread.Sleep(sleep);
            }


            // Mobile searches (the code generates error if not availbe)
            try
            {
                if (driver.Url != "https://account.microsoft.com/rewards/")
                    driver.Navigate().GoToUrl("https://account.microsoft.com/rewards/");

                long maxPointMobile = (long)js.ExecuteScript("return dashboard.userStatus.counters.mobileSearch[0].pointProgressMax");
                pointPC += (long)js.ExecuteScript("return dashboard.userStatus.counters.pcSearch[1].pointProgress");
                long pointMobile = (long)js.ExecuteScript("return dashboard.userStatus.counters.mobileSearch[0].pointProgress");
                long pointMobileLeft = maxPointMobile - pointMobile;

                driver.Quit();
                if (pointMobileLeft == 0 && closeBrowserAtEnd)
                    return;

                //open browser with different user agent to emulate mobile searches
                options.AddArgument("--user-agent=Mozilla/5.0 (Linux; Android 10; Pixel 4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.101 Mobile Safari/537.36");
                driver = new EdgeDriver(options);
                for (long i = 0; i < pointMobileLeft; i += pointForSearch)
                {
                    driver.Navigate().GoToUrl("https://bing.com/search?q=" + rsg.GenString(length));
                    Thread.Sleep(sleep);
                }

                options.AddExcludedArgument("--user-agent");
                driver.Quit();


            }
            catch (WebDriverException e)
            {
                Console.WriteLine("mobile researches not enabled on this level - " + e);
            }
            //not reopen the browser if not requested
            if (!closeBrowserAtEnd)
            {
                driver = new EdgeDriver(options);
            }

        }

        //if not logged click on button login (you must have been logged before, so that after have clicked the button the site will not request account credentials)
        private static void Login()
        {
            try
            {
                driver.FindElement(By.XPath("//a[@id='signinlinkhero']")).Click();
            }
            catch (NoSuchElementException)
            {

            }

        }

        private static string GetEdgeVersion()
        {
            //get version using powershell
            var proc = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = "(Get-AppxPackage -Name \"Microsoft.MicrosoftEdge.Stable\").Version",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                }
            };
            proc.Start();
            //get return value
            string ret = proc.StandardOutput.ReadToEnd();
            //remove \n\r
            return ret.Substring(0, ret.Length - 2);
        }

        private static void DownloadDriver(string path)
        {
            string version = GetEdgeVersion();
            string req = "https://msedgedriver.azureedge.net/" + version + "/edgedriver_win64.zip";
            string zipPath = Path.GetFullPath(path + "edgedriver_win64.zip");
            string exePath = Path.GetFullPath(path + "msedgedriver.exe");


            if (File.Exists(zipPath))
            {
                Console.WriteLine(zipPath + " already exists, please remove it and run again the program");
                //return;
                Application.Exit();
            }

            //System.UnauthorizedAccessException
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

                foreach (ZipArchiveEntry entry in archive.Entries.Where(e => e.FullName == "msedgedriver.exe"))
                {
                    entry.ExtractToFile(exePath);
                }
                //var a = archive.Entries.Where(e => e.FullName == "msedgedriver.exe").FirstOrDefault();
            }
            Console.WriteLine("Removing the zip");
            File.Delete(zipPath);
            Console.WriteLine("Download completed");
        }

        public static void Main(string[] args)
        {
            //manage arguments
            string edgeUser, path;
            try
            {
                Tuple<string, string> paramsRet = Arguments(args);
                edgeUser = paramsRet.Item1;
                path = paramsRet.Item2;
            }
            catch (ProfileNotFound e)
            {
                MessageBox.Show("Check if the selected profile exists", e.Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            options = new EdgeOptions
            {
                UseChromium = true
            };

            //set the profile to use
            options.AddArgument("user-data-dir=" + Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Microsoft\\Edge\\User Data");
            options.AddArguments("profile-directory=" + edgeUser);

            // Create an Edge session, if Edge is already opened or an exception occoured during creation of driver a messagebox will show a message
            try
            {
                driver = new EdgeDriver(path, options);
            }
            catch (WebDriverException e)
            {
                if (File.Exists("msedgedriver.exe"))
                {
                    Console.WriteLine(e);
                    MessageBox.Show("Try to close all Microsoft Edge sessions and to restart the programma", "Impossible to open Edge with the selected profile", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    Console.WriteLine(e);
                    DownloadDriver(path);
                    MessageBox.Show("The program will restart, if it will loop close the app", "The new driver has been installed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Application.Restart();
                    return;
                }
            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine(e);
                DownloadDriver(path);
                MessageBox.Show("The program will restart, if it will loop close the app", "The new driver has been installed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Application.Restart();
                return;

            }
            wait = new WebDriverWait(driver, new TimeSpan(0, 0, 5));

            //go to the rewards home page
            driver.Url = "https://account.microsoft.com/rewards/";

            Login();

            bool doPunchCard = false;

            //try to click pause button in the punchCards section
            //TODO check if it is still needed
            /*
            try
            {
                Click(driver.FindElement(By.XPath("//button[@class='c-action-toggle c-glyph f-toggle glyph-pause']")));
                doPunchCard = true;
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("punch card will be not resolved");
            }*/

            //the code is executed twice so if one or more cards have been missed they can be resolved again
            Console.WriteLine("Starting cards");
            for (int i = 0; i < 2; i++)
            {
                DailyCards();
                OtherCards();
                if (doPunchCard)
                {
                    PunchCard();
                }
            }
            BingResearches(3, true);
            Console.ReadKey();
        }
    }
}
