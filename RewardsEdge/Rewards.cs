using Microsoft.Edge.SeleniumTools;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.Collections.ObjectModel;
using System.Diagnostics;

namespace RewardsEdge
{


    /**
     * <summary> Where is located the entire code of automation of Microsoft Rewards. </summary>
     */
    class Rewards
    {
        private static WebDriverWait wait;
        private static IWebDriver driver;
        private static EdgeOptions options;

        

        /**
         * <summary> Takes and complete the daily cards in the page. </summary>
         */
        private static void DailyCards()
        {
            // get the 3 daily cards
            var listDaily = driver.FindElements(By.XPath("//div[@id='daily-sets']//div[@class='c-card-content']"));

            for (int i = 0; i < 3; i++)
            {
                if (!IsCardDone(listDaily[i]))
                {
                    // click on the card
                    Click(listDaily[i].FindElement(By.XPath(".//div[@class='actionLink x-hidden-vp1']/span")));
                    driver.SwitchTo().Window(driver.WindowHandles.Last());
                    // resolve quiz if exist
                    ResolvePromotion();
                    driver.SwitchTo().Window(driver.WindowHandles[0]);

                }


            }
        }

        /**
         * <summary> Takes and complete the other cards.</summary>
         */
        private static void OtherCards()
        {
            // get the other cards
            var listOthers = driver.FindElements(By.XPath("//mee-card-group[@id='more-activities']//div[@class='c-card-content']"));
            foreach (var other in listOthers)
            {
                if (!IsCardDone(other))
                {
                    // click on the card
                    Click(other.FindElement(By.XPath(".//div[@class='actionLink x-hidden-vp1']")));
                    driver.SwitchTo().Window(driver.WindowHandles.Last());
                    // resolve quiz if exist
                    ResolvePromotion();
                    driver.SwitchTo().Window(driver.WindowHandles[0]);
                }


            }
        }

        /**
         * <summary> Completes the passed punch card. </summary>
         * If a punch card requires to buy something it will be ignored.
         * <param name="punchcard"> The punch card to complete. </param>
         * <param name="sleep"> Time to wait after completed a promotion. </param>
         */
        private static void ResolvePunchCard(IWebElement punchcard, int sleep)
        {
            Console.WriteLine("In punch card");
            // open the punch card page
            //punchcard.FindElement(By.XPath(".//span[@class='pointLink ng-binding ng-scope']")).Click();
            Click(punchcard.FindElement(By.XPath(".//span[@class='pointLink ng-binding ng-scope']")));
            driver.SwitchTo().Window(driver.WindowHandles.Last());
            // for each quiz in the card
            foreach (var toClick in driver.FindElements(By.XPath("//button[@class='btn-primary btn win-color-border-0 card-button-height pull-left margin-right-24 padding-left-24 padding-right-24']/preceding::a[1]")))
            {
                // if the button will not redirect to a bing page stop the program
                string stringUrl = toClick.GetAttribute("href");
                if (!stringUrl.Contains("https%3A%2F%2Faka.ms") && !stringUrl.Contains("https%3A%2F%2Fwww.bing.com"))
                {
                    Console.WriteLine("exit punchcard");
                    break;
                }

                Console.WriteLine("in punchcard quiz");
                Thread.Sleep(sleep);
                // open the url in another page
                Click(toClick);
                driver.SwitchTo().Window(driver.WindowHandles.Last());
                ResolvePromotion();
                driver.SwitchTo().Window(driver.WindowHandles.Last());
            }
            // close page and return to the rewards page
            driver.Close();
            driver.SwitchTo().Window(driver.WindowHandles[0]);
        }

        /**
         * <summary> Completes the punch cards in the page. </summary>
         * to complete a single punch card the method uses <see cref="ResolvePunchCard(IWebElement, int)">ResolvePunchCard</see>.
         * <param name="sleep"> Time to wait after completed a promotion. </param>
         */
        private static void PunchCard(int sleep = 1000)
        {
            Console.WriteLine("In punchcard");
            IWebElement punchCard;
            // try to take the punch card
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
            // for each button there is a punch card
            foreach (var button in punchCard.FindElements(By.XPath("//mee-carousel/div/div[1]/div/button")))
            {
                Click(button);
                IWebElement section = sections[i_section++];

                Console.WriteLine("Name: " + section.FindElement(By.XPath(".//p[@class='c-subheading ng-binding']")).Text);
                bool completed = true;
                foreach (var check in section.FindElements(By.XPath(".//div[@class='icon-container ng-scope']/span/span")))
                {
                    if (check.GetAttribute("class") != "mee-icon mee-icon-StatusCircleOuter checkmark ng-scope")
                    {
                        completed = false;
                        break;
                    }
                }
                // if punch card isn't completed
                if (!completed)
                {
                    ResolvePunchCard(section, sleep);
                }

            }

        }

        /**
         * <summary> Check if card is completed. </summary>
         * <param name="card"> The card to be checked. </param>
         */
        private static bool IsCardDone(IWebElement card)
        {
            return card.FindElements(By.XPath(".//span[@class='mee-icon mee-icon-SkypeCircleCheck']")).Count > 0;
        }

        /**
         * <summary> Click the webElement when possible. </summary>
         * It's used to avoid errors when trying to click. There is a timeout set to 5 seconds
         * <param name="we"> The webElement to click. </param>
         */
        private static void Click(IWebElement we)
        {
            try
            {
                wait.Until(e => we.Displayed && we.Enabled ? we : null);
                we.Click();
            }
            catch (WebDriverTimeoutException e)
            {
                Console.WriteLine("TIMEOUT ERROR: "+ e +"\nElement not found: "+we.ToString()+"\nPress a key to terminate the program");
                driver.Quit();
                Console.ReadKey();
                Environment.Exit(-1);
            }
        }

        /**
         * <summary> Click on a pool option to complete the promotion. </summary>
         * <param name="poolOptions"> All the pool options</param>
         */
        private static void ResolvePool(ReadOnlyCollection<IWebElement> poolOptions)
        {
            Click(poolOptions[0].FindElement(By.XPath(".//div[@id='btoption0']")));
            Thread.Sleep(500);
            driver.Close();
        }

        /**
         * <summary> Completes the promotion. </summary>
         * Check which type of promotion it is and execute the correct function to resolve it (<see cref="ResolvePool(ReadOnlyCollection{IWebElement})">ResolvePool</see> and <see cref="DoQuiz(int)">DoQuiz</see>).
         * It's recommended to set an high sleep time because the quiz overlay not appear immediatly and becauese in case it is a simple resarch if the time the browser stay in the page is short, it may not award points 
         * <param name="sleep"> Time to wait before checking which type it is</param>
         */
        private static void ResolvePromotion(int sleep = 3500)
        {
            // wait before analyze page, so page has the time to load and in case it isn't a quiz the time to get the card as completed
            Thread.Sleep(sleep);
            // If there is no quiz exit from function
            var overlay = driver.FindElements(By.XPath("//div[@class='btOverlay']"));
            if (overlay.Count == 0)
            {
                driver.Close();
                return;
            }

            // if there is a pool
            var findPoll = overlay[0].FindElements(By.XPath(".//div[@class='bt_poll']"));
            if (findPoll.Count > 0)
            {
                ResolvePool(findPoll);
                return;
            }

            // if there is a quiz
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

        /**
         * <summary> Recognizes which quiz it is and call the right function to complete it.</summary>
         * Supported quiz: <see cref="MultipleAnswerQuiz(int)">multiple answer quiz</see>, <see cref="ThisOrThat(int)">this or that</see> and <see cref="SingleAnswerQuiz(int)">single answer quiz</see>.
         * <param name="sleep"> Sleep parameter to pass to the other functions. </param>
         */
        private static void DoQuiz(int sleep = 3500)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            if ((bool)js.ExecuteScript("return _w.rewardsQuizRenderInfo.isListicleQuizType"))
            {
                MultipleAnswerQuiz(sleep);
            }
            else if ((bool)js.ExecuteScript("return _w.rewardsQuizRenderInfo.isWOTQuizType"))
            {
                ThisOrThat();
            }
            else
            {
                SingleAnswerQuiz(sleep);
            }

        }

        /**
         * <summary> Complete the multiple answer quiz</summary>
         * <param name="sleep"> Time to wait after have been completed a sub-quiz</param>
         */
        private static void MultipleAnswerQuiz(int sleep)
        {
            var maxAndCurrent = getMaxAndCurrent();
            Console.WriteLine("current: " + maxAndCurrent.Current + ", max: " + maxAndCurrent.Max);
            for (long i = maxAndCurrent.Current; i <= maxAndCurrent.Max; i++)
            {
                // 8 possible answers
                for (int j = 0; j < 8; j++)
                {
                    var slide = driver.FindElement(By.XPath("//div[@class='btOverlay']//div[@class='slide']/div[@id]"));
                    Click(slide);
                    Thread.Sleep(100);
                    // when all the correct answers have been selected this div tag with this class will appear
                    var temp = driver.FindElements(By.XPath("//div[@class='btOverlay']//div[@class='b_promtxt rqQPanel b_hide']"));
                    if (temp.Count > 0)
                        break;
                    // if this is the last sub-quiz there is another way to recognize if quiz is ended
                    else if (i == maxAndCurrent.Max)
                    {
                        var temp2 = driver.FindElements(By.XPath("//div[@class='btOverlay']//div[@class='headerMessage']"));
                        if (temp2.Count > 0)
                            break;
                    }
                }
                Thread.Sleep(sleep);
            }
        }

        /**
         * <summary> Complete the single answer quiz</summary>
         * <param name="sleep"> Time to wait after have been completed a sub-quiz</param>
         */
        private static void SingleAnswerQuiz(int sleep)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            var maxAndCurrent = getMaxAndCurrent(js);
            long numOptions = (long)js.ExecuteScript("return _w.rewardsQuizRenderInfo.numberOfOptions");
            for (long i = maxAndCurrent.Current; i <= maxAndCurrent.Max; i++)
            {
                // 4 possible answers
                for (long j = 0; j < numOptions; j++)
                {
                    var slide = driver.FindElement(By.XPath("//div[@class='btOverlay']//input[@class='rqOption']"));
                    Click(slide);
                    Thread.Sleep(100);
                    // check if i selected the correct answer
                    if (driver.FindElements(By.XPath("//div[@class='btOverlay']//input[@class='rqOption correctAnswer']")).Count > 0)
                        break;
                }
                Thread.Sleep(sleep);
            }

        }

        /**
         * <summary> Complete the "this or that" quiz</summary>
         */
        private static void ThisOrThat()
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            var maxAndCurrent = getMaxAndCurrent(js);
            for (long i = maxAndCurrent.Current; i <= maxAndCurrent.Max; i++)
            {
                // this value will be used to calculate which is the correct answer
                string IG = driver.FindElement(By.XPath("//span[@id='nc_iid']")).GetAttribute("_ig");
                // get the correct answer value
                string CorrectAnswer = (string)js.ExecuteScript("return _w.rewardsQuizRenderInfo.correctAnswer");

                var Options = driver.FindElements(By.XPath("//div[@class='btOverlay']//div[@class='btOptionCard' and @id]"));
                // calculate value of first answer
                string firstOptValue = ResolveCorrectAnswer(Options[0].GetAttribute("data-option"), IG);

                if (CorrectAnswer == firstOptValue)
                    Click(Options[0]);
                else
                    Click(Options[1]);
            }
        }

        /**
         * <summary> Get the maximus and the current points in a quiz.</summary>
         * <returns> Maximus and the current points in a quiz.</returns>
         */
        private static (long Max, long Current) getMaxAndCurrent()
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            return getMaxAndCurrent(js);
        }

        /**
         * <summary> Get the maximum and the current points in a quiz with the passed js executor.</summary>
         * <param name="js"> The js executor to use</param>
         * <returns> Maximum and the current points in a quiz.</returns>
         */
        private static (long Max, long Current) getMaxAndCurrent(IJavaScriptExecutor js)
        {
            long max = (long)js.ExecuteScript("return _w.rewardsQuizRenderInfo.maxQuestions");
            long curr = (long)js.ExecuteScript("return _w.rewardsQuizRenderInfo.currentQuestionNumber");
            return (Max: max, Current: curr);
        }

        // TODO complete documentation
        /**
         * <summary> Gets the correct answer value in this or that quiz.</summary>
         * Trasposition in C# of function "br" present in the html of this or that quiz.
         * <param name="dataOption"> The data option. </param>
         * <param name="IG"> IG value. </param>
         * <returns> The correct answer value.</returns>
         */
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

        /**
         * <summary> Makes PC and mobile Bing researches to earn points.</summary>
         * <param name="pointForSearch"> How many points are eanred for research.</param>
         * <param name="closeBrowserAtEnd"> If browser needs to be closed and the program ends.</param>
         * <param name="length"> How long the research string must be.</param>
         * <param name="sleep"> Time to wait after a research</param>
         */
        private static void BingResearches(long pointForSearch, bool closeBrowserAtEnd = false, int length = 4, int sleep = 1000)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            // get how many searches remain for desktop
            long maxPointPC = (long)js.ExecuteScript("return dashboard.userStatus.counters.pcSearch[0].pointProgressMax");
            maxPointPC += (long)js.ExecuteScript("return dashboard.userStatus.counters.pcSearch[1].pointProgressMax");
            long pointPC = (long)js.ExecuteScript("return dashboard.userStatus.counters.pcSearch[0].pointProgress");
            try
            {
                // get how many searches remain for mobile
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

                // open browser with different user agent to emulate mobile searches
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
            // not reopen the browser if not requested
            if (!closeBrowserAtEnd)
            {
                driver = new EdgeDriver(options);
            }

        }

        /**
         * <summary> If not logged click on button login.</summary>
         * you must have been logged before, so that after have clicked the button the site will not request account credentials.
         */
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

        public void Close()
        {

        }

        public static void Main(string[] args)
        {
            // manage arguments
            string profileFolder, path, userDataDir;
            try
            {
                Tuple<string, string, string> paramsRet = EdgeManagment.Arguments(args);
                profileFolder = paramsRet.Item1;
                userDataDir = paramsRet.Item2;
                path = paramsRet.Item3;
            }
            catch (ProfileNotFound e)
            {
                MessageBox.Show("Check if the selected profile exists", e.Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // set Edge chromium
            options = new EdgeOptions
            {
                UseChromium = true
            };

            //set the profile to use
            options.AddArgument("user-data-dir=" + userDataDir);
            options.AddArguments("profile-directory=" + profileFolder);

            // Create an Edge session
            try
            {
                driver = new EdgeDriver(path, options);
            }
            catch (WebDriverException e)
            {
                // probably the error is caused by an already open session of edge with the selected profile
                if (File.Exists("msedgedriver.exe"))
                {
                    Console.WriteLine("ERROR: "+e);
                    MessageBox.Show("Try to close all Microsoft Edge sessions and to restart the programma", "Impossible to open Edge with the selected profile", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                // install the driver
                else
                {
                    Console.WriteLine("ERROR: " + e);
                    EdgeManagment.DownloadDriver(path);
                    MessageBox.Show("The program will restart, if it will loop close the app", "The new driver has been installed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Application.Restart();
                    return;
                }
            }
            // raised when driver can't be used with the installed edge version, the correct driver version will be downloaded
            catch (InvalidOperationException e)
            {
                Console.WriteLine("ERROR: " + e);
                EdgeManagment.DownloadDriver(path);
                MessageBox.Show("The program will restart, if it will loop close the app", "The new driver has been installed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Application.Restart();
                return;

            }
            wait = new WebDriverWait(driver, new TimeSpan(0, 0, 5));

            // go to the rewards home page
            driver.Url = "https://account.microsoft.com/rewards/";

            Login();
            Click(driver.FindElement(By.XPath("//script")));
            bool doPunchCard = true;
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
            //Console.ReadKey();
        }
    }
}
