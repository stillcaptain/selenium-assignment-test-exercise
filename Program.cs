using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;

namespace Test_Exercise_1
{
    class Program
    {
        static void Main(string[] args)
        {
            IWebDriver driver = new FirefoxDriver();
            driver.Manage().Window.Maximize();

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(25));

            try
            {
                driver.Navigate().GoToUrl("https://go.microsoft.com/fwlink/p/?linkid=2125442&clcid=0x409&culture=en-us&country=us");

                // Email
                IWebElement emailInput = wait.Until(delegate (IWebDriver driver)
                {
                    IWebElement element = driver.FindElement(By.Name("loginfmt"));
                    return element;
                });
                emailInput.SendKeys("USERNAME" + Keys.Enter);

                // Wait for password input
                IWebElement passwordInput = wait.Until(delegate (IWebDriver driver)
                {
                    IWebElement element = driver.FindElement(By.Id("passwordEntry"));
                    return element;
                });

                //input password credential
                passwordInput.SendKeys("PASSWORD");
                passwordInput.SendKeys(Keys.Enter);


                // refuse at "Stay signed in?" prompt
                try
                {
                    IWebElement staySignedIn = wait.Until(delegate (IWebDriver driver)
                    {
                        IWebElement element = driver.FindElement(By.CssSelector("button[data-testid='secondaryButton'][type='submit']"));

                        return element;
                    });

                    staySignedIn.Click();
                }
                catch
                {
                    Console.WriteLine("No 'Stay signed in' prompt appeared.");
                }

                Console.WriteLine("Login attempted. Waiting for inbox...");


                // Wait until Focused tab exists
                By focusedTabLocator = By.CssSelector("button[role='tab'][name='Focused']");
                wait.Until(delegate (IWebDriver driver)
                {
                    return driver.FindElements(focusedTabLocator).Count > 0;
                });

                IWebElement focusedTab = wait.Until(delegate (IWebDriver driver)
                {
                    IWebElement element = driver.FindElement(focusedTabLocator);
                    return element;
                });

                string ariaSelected = focusedTab.GetAttribute("aria-selected");

                if (ariaSelected == "true")
                {
                    Console.WriteLine("Focused tab is already selected.");
                }
                else
                {
                    Console.WriteLine("Focused tab not selected. Clicking it...");
                    SafeClick(driver, wait, focusedTabLocator);
                }


                //Selectors for mail row and checkbox
                By mailRows = By.CssSelector("div[role='option'][data-focusable-row='true']");
                By rowCheckbox = By.CssSelector("div[role='checkbox'][aria-label='Select a conversation']");
                By itemsSelectedText = By.XPath("//*[contains(., 'items selected')]");

                // Wait for at least 2 rows to exist
                wait.Until(delegate (IWebDriver driver)
                {
                    return driver.FindElements(mailRows).Count >= 2;
                });

                var rows = driver.FindElements(mailRows);

                // Click checkbox in row 1 and row 2
                ClickRowCheckbox(driver, wait, rows[0]);
                ClickRowCheckbox(driver, wait, rows[1]);

                // Verify "2 items selected"
                wait.Until(delegate (IWebDriver driver)
                {
                    return driver.PageSource.Contains("2 conversations selected");
                });
                Console.WriteLine("Verified: 2 conversations selected.");



                //Get sender and subject of first n mails 
                int n = 1;
                var (sender, subject) = GetSenderAndSubject(rows[n - 1]);

                Console.WriteLine($"Email #{n}: Sender='{sender}', Subject='{subject}'");



                Console.WriteLine("Press ENTER to close.");
                Console.ReadLine();

            }
            finally
            {
                driver.Quit();
            }
        }


        //Functions

        //Safe Click for tab selection function
        static void SafeClick(IWebDriver driver, WebDriverWait wait, By elementLocator)
        {
            // Refetch element right before clicking
            IWebElement element = wait.Until(delegate (IWebDriver driver)
            {
                IWebElement foundElement = driver.FindElement(elementLocator);
                return foundElement;
            });

            // Scroll into view so headers/containers are less likely to overlap
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({block:'center', inline:'center'});", element);

            try
            {
                element.Click();
            }
            catch (ElementClickInterceptedException)
            {
                // Fallback: JS click bypasses the “container obscures it” warning in terminal
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", element);
            }
            catch (StaleElementReferenceException)
            {
                element = wait.Until(delegate (IWebDriver driver)
                {
                    IWebElement foundElement = driver.FindElement(elementLocator);
                    return foundElement;
                });

                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", element);
            }
        }


        // Row Checkbox Click function
        static void ClickRowCheckbox(IWebDriver driver, WebDriverWait wait, IWebElement row)
        {
            // Find the checkbox *within that row*
            IWebElement cb = row.FindElement(By.CssSelector("div[role='checkbox'][aria-label='Select a conversation']"));

            // Scroll the checkbox into view
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({block:'center', inline:'center'});", cb);

            // Try normal click, fallback to JS click if intercepted
            try
            {
                cb.Click();
            }
            catch (ElementClickInterceptedException)
            {
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", cb);
            }
            catch (StaleElementReferenceException)
            {
                // If the row re-rendered, re-find from the row again
                cb = row.FindElement(By.CssSelector("div[role='checkbox'][aria-label='Select a conversation']"));
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", cb);
            }

        }
        

        //Get sender and subject function
        static (string Sender, string Subject) GetSenderAndSubject(IWebElement row)
        {
            // Sender is first span with a title attribute inside the sender area. More stable than matching CSS classes.
            var senderSpan = row.FindElement(By.CssSelector("div.ESO13 span[title]"));
            string sender = senderSpan.Text.Trim();

            // Subject span
            var subjectSpan = row.FindElement(By.CssSelector("span.TtcXM"));
            string subject = subjectSpan.Text.Trim();

            return (sender, subject);
        }


    }
}