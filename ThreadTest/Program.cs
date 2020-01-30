using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using Cookie = OpenQA.Selenium.Cookie;
using System.Threading.Tasks;

namespace ClaroTests
{
    class Test
    {
        static void Main(string[] args)
        {

            var t1 = new Task(Thread1);
            t1.Start();

            var t2 = new Task(Thread2);
            t2.Start();

            var t3 = new Task(Thread3);
            t3.Start();

            var t4 = new Task(Thread4);
            t4.Start();

            Task.WaitAll(t1, t2, t3, t4);
        }

        public static void Thread1()
        {
            DefaultThread(1);
        }

        public static void Thread2()
        {
            DefaultThread(2);
        }

        public static void Thread3()
        {
            DefaultThread(3);
        }

        public static void Thread4()
        {
            DefaultThread(4);
        }

        public static void DefaultThread(int y)
        {
            SqlConnection sqlConnection = new SqlConnection("Data Source=DESK-INOVA-0029;Initial Catalog=ClaroDB;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            sqlConnection.Open();
            SqlDataAdapter adapter = new SqlDataAdapter();

            ChromeOptions options = new ChromeOptions();
            options.AddArguments("--headless");

            IWebDriver driver = new ChromeDriver(options);
            driver.Url = ("https://appclarotvprepago.visiontec.com.br/Account");
            Wait(1000);

            var login = driver.FindElement(By.Id("Login"));
            login.SendKeys("82211409024");
            Wait(1000);

            var pw = driver.FindElement(By.Id("Password"));
            pw.SendKeys("567890");
            Wait(1000);

            var submit = driver.FindElement(By.ClassName("btn"));
            submit.Click();
            Wait(1000);

            var find = driver.FindElement(By.XPath("//*[@id='accordion']/div/div/div[1]/a"));
            find.Click();

            Wait(1000);

            String cepNumber;
            StreamReader sr = new StreamReader($"C:\\CEPs-{y}.txt");

            cepNumber = sr.ReadLine();
            var cep = driver.FindElement(By.Id("CEP"));
            foreach (var number in cepNumber)
            {
                cep.SendKeys($"{number}");
                Wait(250);
            }
            var getCep = driver.FindElement(By.XPath("//*[@id='CepForm']/div/div/div/button"));
            getCep.Click();
            var loaded = driver.FindElement(By.ClassName("Loading"));
            WaitUntilLoad(loaded);

            IEnumerable<Cookie> cookies = driver.Manage().Cookies.AllCookies;

            var cookieContainer = new CookieContainer();

            foreach (Cookie ck in cookies)
            {
                cookieContainer.Add(new System.Net.Cookie(ck.Name, ck.Value, ck.Path, ck.Domain));
            }

            while (cepNumber != null)
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://appclarotvprepago.visiontec.com.br/SearchInstaller/SearchAsync");

                var postData = $"cep={cepNumber}";
                var data = Encoding.ASCII.GetBytes(postData);
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = data.Length;
                request.CookieContainer = cookieContainer;
                string html;

                using (var stream = request.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                }

                WebResponse response;

                try
                {
                    response = request.GetResponse();
                }
                catch (WebException)
                {
                    continue;
                }

                using (Stream dataStream = response.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(dataStream);

                    html = reader.ReadToEnd();
                    request.Abort();
                    response.Close();
                }

                var hDoc = new HtmlDocument();
                html = HtmlEntity.DeEntitize(html);
                hDoc.LoadHtml(html);
                var root = hDoc.DocumentNode;

                string howManyText = "";
                try
                {
                    howManyText = root.Descendants()
                    .Where(n => n.GetAttributeValue("class", "").Equals("card"))
                    .Single()
                    .Descendants("h5")
                    .Single()
                    .InnerText;
                }
                catch (InvalidOperationException e)
                {
                    try
                    {
                        var sql = $"Insert into Cep (CepNumber, TecNumber) values ('{cepNumber}', '0')";

                        adapter.InsertCommand = new SqlCommand(sql, sqlConnection);
                        adapter.InsertCommand.ExecuteNonQuery();

                        adapter.InsertCommand.Dispose();
                        cepNumber = sr.ReadLine();
                        continue;
                    }
                    catch (SqlException _e)
                    {

                    }
                };

                string howManyNumber = string.Empty;
                int howMany = 0;
                for (int i = 0; i < howManyText.Length; i++)
                {
                    if (Char.IsDigit(howManyText[i]))
                        howManyNumber += howManyText[i];
                }

                if (howManyNumber.Length > 0)
                    howMany = int.Parse(howManyNumber);

                for (int i = 0; i < howMany; i++)
                {
                    var tecId = i + 1;
                    var tecName = root.Descendants()
                        .Where(n => n.GetAttributeValue("href", "").Equals($"#installer_{i}"))
                        .Single()
                        .InnerText
                        .Trim();

                    tecName = RemoveDiacritics(tecName);

                    var tecInfoNumber = root.Descendants()
                        .Where(n => n.GetAttributeValue("id", "").Equals($"installer_{i}"))
                        .Single()
                        .Descendants()
                        .Where(n => n.GetAttributeValue("class", "").Equals("card-body"))
                        .Single()
                        .Descendants("p")
                        .Count();

                    var tecPhones = root.Descendants()
                        .Where(n => n.GetAttributeValue("id", "").Equals($"installer_{i}"))
                        .Single()
                        .Descendants()
                        .Where(n => n.GetAttributeValue("class", "").Equals("card-body"))
                        .Single()
                        .Descendants("p")
                        .Where(n => n.ChildNodes.Count == 1 && !(n.InnerText.Contains("@")));

                    var tecEmail = root.Descendants()
                        .Where(n => n.GetAttributeValue("id", "").Equals($"installer_{i}"))
                        .Single()
                        .Descendants()
                        .Where(n => n.GetAttributeValue("class", "").Equals("card-body"))
                        .Single()
                        .Descendants("p")
                        .Last()
                        .InnerText;

                    string tecPhone1 = "";
                    string tecPhone2 = "";
                    string tecPhone3 = "";
                    string tecPhone4 = "";

                    switch (tecPhones.Count())
                    {
                        case 1:
                            tecPhone1 = tecPhones.FirstOrDefault().InnerText;
                            break;

                        case 2:
                            tecPhone1 = tecPhones.FirstOrDefault().InnerText;
                            tecPhone2 = tecPhones.LastOrDefault().InnerText;
                            break;

                        case 3:
                            tecPhone1 = tecPhones.FirstOrDefault().FirstChild.InnerText;
                            tecPhone2 = tecPhones.First().NextSibling.NextSibling.InnerText;
                            tecPhone3 = tecPhones.LastOrDefault().InnerText;
                            break;

                        case 4:
                            tecPhone1 = tecPhones.FirstOrDefault().InnerText;
                            tecPhone2 = tecPhones.First().NextSibling.NextSibling.InnerText;
                            tecPhone3 = tecPhones.Last().PreviousSibling.PreviousSibling.InnerText;
                            tecPhone4 = tecPhones.LastOrDefault().InnerText;
                            break;
                    }
                    try
                    {
                        var sql = $"Insert into Technical (TecId, TecName, TecPhone1, TecPhone2, TecPhone3, TecPhone4, TecEmail) " +
                        $"values ('{tecId}','{tecName}','{tecPhone1}','{tecPhone2}','{tecPhone3}','{tecPhone4}','{tecEmail}') " +
                        $"Insert into Cep(CepNumber, TecNumber) values('{cepNumber}', '{howMany}') " +
                         $"Insert into CepTechnical(Cep, Technical) values('{cepNumber}', '{tecName}')";

                        adapter.InsertCommand = new SqlCommand(sql, sqlConnection);
                        adapter.InsertCommand.ExecuteNonQuery();

                        adapter.InsertCommand.Dispose();
                    }
                    catch (SqlException Ex)
                    {
                        if (Ex.Number == 2627)
                        {

                        }
                    }
                }

                cep.Clear();
                cepNumber = sr.ReadLine();
            }
            sqlConnection.Close();
            driver.Quit();
        }

        static void Wait(int milsec)
        {
            Thread.Sleep(milsec / 2);
        }

        static void WaitUntilLoad(IWebElement loaded)
        {
            var loadingValue = loaded.GetCssValue("display");

            while (loadingValue == "block")
            {
                loadingValue = loaded.GetCssValue("display");
            }
        }

        static string RemoveDiacritics(string text)
        {
            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();

            foreach (var c in normalizedString)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
        }
    }
}