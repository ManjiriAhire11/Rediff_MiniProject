using System;
using Aspose.Cells;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using System.Collections.Generic;
using System.IO;

namespace MiniProject
{
    public class MoneyRediffProject
    {
        public static List<string> output = new List<string>();
        static void Main(string[] args)
        {

            output.Add("Program Started");

            // Launch Chrome
             IWebDriver driver = new ChromeDriver(@"C:\Users\manji\source\repos\MPro");
            // Launch Firefox
            //IWebDriver driver = new FirefoxDriver(@"C:\Users\manji\source\repos\MPro");

            // Login to money.rediff.com
            Login(driver);
            System.Threading.Thread.Sleep(2000);

            Worksheet sheet = ReadExcel();
            // Loop through cells data
            for (int row = 0; row <= sheet.Cells.MaxDataRow; row++)
            {
                // Thread Sleep
                System.Threading.Thread.Sleep(2000);

                // Stock Data Read From Excel Sheet - Line By Line
                string name = sheet.Cells[row, 0].Value.ToString(); // Tech Mahindra
                string dateOfPurchase = sheet.Cells[row, 1].Value.ToString(); // 22-09-2021 00:00:00
                string quantity = sheet.Cells[row, 2].Value.ToString(); // 100
                string purchasePrice = sheet.Cells[row, 3].Value.ToString(); // 250
                string exchange = sheet.Cells[row, 4].Value.ToString(); // BSE

                // Check If Stock Exists in Table
                var isExists = CheckStockExistancy(driver, name);
                {
                    Console.WriteLine("Portfolio is exist");
                }

                // If Stock does not Exists add that Stock
                if (!isExists)
                {
                    dateOfPurchase = dateOfPurchase.Split(' ')[0];
                    AddStock(driver, name, dateOfPurchase, quantity, purchasePrice, exchange);
                    
                }
            }

            output.Add("Program Ended");

            writeOutputFile();
        }

        public static void Login(IWebDriver driver)
        {
            // Read Config File
            string[] configLines = System.IO.File.ReadAllLines(@"C:\Users\manji\source\repos\MPro\config.txt");
            // Maximize the browser
            driver.Manage().Window.Maximize();

            // Launch Website money.rediff.com
            driver.Url = configLines[0];

            // Get signIn page link
            IWebElement signInPage = driver.FindElement(By.XPath("//*[@id='signin_info']/a[1]"));
            // Click signIn page link
            signInPage.Click();

            // Email Input
            IWebElement emailInput = driver.FindElement(By.XPath("//*[@id='useremail']"));
            // Fill input with email
            emailInput.SendKeys(configLines[1]);

            // Password Input
            IWebElement passwordInput = driver.FindElement(By.XPath("//*[@id='userpass']"));
            // Fill input with email
            passwordInput.SendKeys(configLines[2]);

            // Submit Button (Login Button)
            IWebElement submitButton = driver.FindElement(By.XPath("//*[@id='loginsubmit']"));
            // Fill input with email
            submitButton.Click();

            output.Add("Login – Pass - Logged In Successfully");
        }

        public static void AddStock(IWebDriver driver, string name, string dateOfPurchase, string quantity, string purchasePrice, string exchange)
        {
            // Add Stock Button
            IWebElement AddStockButton = driver.FindElement(By.XPath("//*[@id='addStock']"));
            // Click Add Stock Button
            AddStockButton.Click();

            // Stock Name
            IWebElement StockName = driver.FindElement(By.XPath("//*[@id='addstockname']"));
            // Fill Stock Name
            StockName.SendKeys(name);
            // Wait till it loads suggestions
            System.Threading.Thread.Sleep(1000);
            // Select first suggestion
            StockName.SendKeys(Keys.Enter);

            // Stock Date Of Purchase
            IWebElement StockDateOfPurchase = driver.FindElement(By.XPath("//*[@id='stockAddDate']"));
            // Fill Date Of Purchase
            StockDateOfPurchase.SendKeys(dateOfPurchase);

            // Stock Quantity
            IWebElement StockQuantity = driver.FindElement(By.XPath("//*[@id='addstockqty']"));
            // Fill Quantity
            StockQuantity.SendKeys(quantity);

            // Stock Purchase Price
            IWebElement StockPurchasePrice = driver.FindElement(By.XPath("//*[@id='addstockprice']"));
            // Fill Purchase Price
            StockPurchasePrice.SendKeys(purchasePrice);

            // Stock Exchange
            IWebElement BSEStockExchange = driver.FindElement(By.XPath("//*[@id='exchange_tab']/span[1]/label"));
            IWebElement NSEStockExchange = driver.FindElement(By.XPath("//*[@id='exchange_tab']/span[3]/label"));
            // Select Stock Exchange
            if (exchange == "BSE")
            {
                BSEStockExchange.Click();
            }
            else
            {
                NSEStockExchange.Click();
            }

            // Submit Button
            IWebElement SubmitButton = driver.FindElement(By.XPath("//*[@id='addStockButton']"));
            // Click Submit Button
            SubmitButton.Click();

            output.Add("Stock - " + name + " – Pass - Added Stock Successfully");
        }

        public static Boolean CheckStockExistancy(IWebDriver driver, string name)
        {
            // Get Stock Table
            var StockTableTrList = driver.FindElements(By.XPath("//*[@id='stock']/tbody/tr"));
            foreach (IWebElement element in StockTableTrList)
            {
                // Checking if the Stock name already exists or not
                if (element.Text.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    output.Add("Stock - " + name + " - Exists Already!!");
                    return true;
                }
            }
            return false;
        }

        public static Worksheet ReadExcel()
        {
            // Path for Input XLSX
            string path = @"C:\Users\manji\source\repos\MPro\input.xlsx";
            // Read Workbook
            Workbook w = new Workbook(path);
            // Get Sheet
            Worksheet sheet = w.Worksheets[0];

            // Get Cells using its row and column
            // Cell cell = sheet.Cells.GetCell(0, 0);
            // string value = cell.Value.ToString();
            output.Add("Read Excel - Pass - Read Excel Successful!");
            return sheet;
        }

        public static void writeOutputFile()
        {
            String[] outputLines = output.ToArray();
            File.WriteAllLines("output.txt", outputLines);
        }
    }
}