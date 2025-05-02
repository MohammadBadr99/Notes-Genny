using System;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Net;
using System.Windows.Forms;
using SeleniumExtras.WaitHelpers;
using System.Collections.Generic;
using System.Diagnostics;
using System.Security.Policy;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Math;
using OfficeOpenXml.Utils;
using OpenQA.Selenium.Interactions;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections;

namespace SSR___Case_Notes_Creator
{
    public partial class Form1 : Form
    {
        // Case Number initiation
        string caseNumber = "";
        string serialNumber = "";
        string phoneNumber = "";
        string problemDesc = "";
        string street1 = "";
        string street2 = "";

        int languageSelection = 0;

        List<string> originalItems = new List<string>();
        ArrayList langs = ["English", "French", "Spanish"];

        // Define the file path to save the Excel file
        string excelFilePath = "Extracted Data.xlsx";

        public static string WaitForDownloadToComplete(string downloadDirectory, TimeSpan timeout)
        {
            DateTime start = DateTime.Now;

            while (DateTime.Now - start < timeout)
            {
                var files = Directory.GetFiles(downloadDirectory, "*.xls*"); // Check for .xls or .xlsx files
                if (files.Length > 0)
                {
                    return files[0]; // Return the first downloaded file path
                }

                System.Threading.Thread.Sleep(1000); // Wait for 1 second before checking again
            }

            throw new Exception("Download did not complete within the specified timeout.");
        }

        private void PopulateCheckboxListUsingFilterFromExcel(string filePath)
        {
            // Clear existing items in the CheckedListBox
            checkedListBox1.Items.Clear();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets[0]; // Assuming data is in the first worksheet
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 1; row <= rowCount; row += 1)
                {
                    var cellValue = worksheet.Cells[row, 1].Text; // Read from column 1
                    var cellValue1 = worksheet.Cells[row, 2].Text;
                    var cellValue2 = worksheet.Cells[row, 3].Text;
                    if (comboBox1.Text == cellValue1)
                    {
                        checkedListBox1.Items.Add(cellValue + " - " + cellValue2 + " - " + cellValue1);
                    }
                }
            }
        }

        private void SearchCheckboxListUsingFilterFromExcel(string filePath, string text)
        {
            // Clear existing items in the CheckedListBox
            checkedListBox1.Items.Clear();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets[0]; // Assuming data is in the first worksheet
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 1; row <= rowCount; row += 1)
                {
                    var cellValue = worksheet.Cells[row, 1].Text; // Read from column 1
                    var cellValue1 = worksheet.Cells[row, 2].Text;
                    var cellValue2 = worksheet.Cells[row, 3].Text;
                    if (cellValue1.ToLower().Contains(text) || cellValue.ToLower().Contains(text) || cellValue2.ToLower().Contains(text))
                    {
                        checkedListBox1.Items.Add(cellValue + " - " + cellValue2 + " - " + cellValue1);
                    }
                }
            }
        }

        private void PopulateDropDownListFromExcel(string filePath)
        {
            // Clear existing items in the CheckedListBox
            comboBox1.Items.Clear();
            comboBox1.Items.Add("Clear Filter");
            comboBox1.SelectedIndex = 0;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets[0]; // Assuming data is in the first worksheet
                int rowCount = worksheet.Dimension.Rows;

                // Read data from the first column
                int column3 = 2;
                int i = 0;
                int flag = 1;
                for (int row = 1; row <= rowCount; row += 1)
                {
                    var cellValue1 = worksheet.Cells[row, column3].Text;
                    if (!string.IsNullOrWhiteSpace(cellValue1))
                    {
                        if (comboBox1.Items.Count != 0)
                        {
                            for (i = 0; i < comboBox1.Items.Count; i++)
                            {
                                if (comboBox1.Items[i] == cellValue1)
                                {
                                    flag = 1;
                                }
                                if (flag == 1)
                                {
                                    flag = 0;
                                    break;
                                }
                                if (i == comboBox1.Items.Count - 1)
                                {
                                    comboBox1.Items.Add(cellValue1);
                                    break;
                                }
                            }
                        }
                        else
                        {
                            comboBox1.Items.Add(cellValue1);
                        }
                    }
                }
                comboBox1.Sorted = true;
            }
        }

        private void PopulateCheckboxListFromExcel(string filePath)
        {
            // Clear existing items in the CheckedListBox
            checkedListBox1.Items.Clear();


            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets[0]; // Assuming data is in the first worksheet
                int rowCount = worksheet.Dimension.Rows;
                //int skipRow = 1;

                // Read data from the first column
                for (int row = 2; row <= rowCount; row += 1)
                {
                    var cellValue = worksheet.Cells[row, 1].Text; // Read from column 1
                    var cellValue1 = worksheet.Cells[row, 2].Text;
                    var cellValue2 = worksheet.Cells[row, 3].Text;
                    if (!string.IsNullOrWhiteSpace(cellValue))
                    {
                        checkedListBox1.Items.Add(cellValue + " - " + cellValue2 + " - " + cellValue1);
                    }
                }
            }
        }

        public Form1()
        {
            InitializeComponent();
            comboBox1.TextChanged += comboBox1_TextChanged;
        }

        private void caseLink_TextChanged(object sender, EventArgs e)
        {

        }

        private void scraperButton_Click(object sender, EventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("Proceed?", "Confirmation", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                checkedListBox1.Items.Clear();
                checkedListBox2.Items.Clear();
                comboBox1.Items.Clear();
                comboBox1.Enabled = false;

                string caseUrl = caseLink.Text;

                LoadingScreen loadingScreen = new LoadingScreen();
                loadingScreen.Show();

                // Get the directory where the application is running
                string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;


                // Define a persistent directory for user data
                string userDataDir = Path.Combine(baseDirectory, "User Data");

                // Set ChromeOptions to use the user data directory
                var dataCredintials = new ChromeOptions();

                ChromeDriverService service = ChromeDriverService.CreateDefaultService();
                service.HideCommandPromptWindow = true; // Hides the CMD window
                string downloadPath = AppDomain.CurrentDomain.BaseDirectory;
                if (loggedIn.Checked)
                {
                    dataCredintials.AddArgument("--headless=new");
                    dataCredintials.AddArgument("--no_sandbox");
                }
                dataCredintials.AddArgument($"--user-data-dir={userDataDir}");
                dataCredintials.AddArgument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
                dataCredintials.AddUserProfilePreference("download.default_directory", downloadPath);
                dataCredintials.AddUserProfilePreference("download.prompt_for_download", false);
                dataCredintials.AddUserProfilePreference("download.directory_upgrade", true);
                dataCredintials.AddUserProfilePreference("safebrowsing.enabled", true);
                dataCredintials.AddArgument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
                dataCredintials.AddArgument("--window-size=1920,1080");

                IWebDriver? driverMSD = new ChromeDriver(service, dataCredintials);

                bool IsButtonVisible(IWebDriver driver, By by)
                {
                    try
                    {
                        return driver.FindElement(by).Displayed; // Only true if visible
                    }
                    catch (NoSuchElementException)
                    {
                        return false;
                    }
                }

                try
                {
                    using (driverMSD)
                    {
                        driverMSD.Navigate().GoToUrl(caseUrl);

                        driverMSD.Manage().Window.FullScreen();

                        if (loggedIn.Checked == false)
                        {
                            Thread.Sleep(1200000); // Wait for 2 mins to allow login (adjust as needed)
                        }

                        Thread.Sleep(1000);

                        if (IsButtonVisible(driverMSD, By.CssSelector("a[name='cannotAccessAccount']")))
                        {

                            loadingScreen.Dispose();
                            service.Dispose();
                            driverMSD.Quit();

                            throw new Exception("Please login first! \n\n Set the login status to *No* and prepare your credentials to login!");
                        }

                        /*Thread.Sleep(3000);

                        IWebElement scrollableDiv = driverMSD.FindElement(By.Id("tab-section0"));

                        IJavaScriptExecutor js = (IJavaScriptExecutor)driverMSD;

                        js.ExecuteScript("arguments[0].scrollTo({ top: arguments[0].scrollHeight, behavior: 'smooth' });", scrollableDiv);

                        Thread.Sleep(2000); // Wait to see the effect

                        js.ExecuteScript("arguments[0].scrollTo({ top: 0, behavior: 'smooth' });", scrollableDiv);

                        // Find the input field using CSS Selector
                        IWebElement street2Input = wait.Until(ExpectedConditions.ElementExists(By.CssSelector("input[aria-label='Street 2']")));

                        // Get the value of the input field
                        street2 = street2Input.GetAttribute("title");

                        // Find the input field using CSS Selector
                        IWebElement street1Input = wait.Until(ExpectedConditions.ElementExists(By.CssSelector("input[aria-label='Street 1']")));

                        // Get the value of the input field
                        street1 = street1Input.GetAttribute("title");*/

                        WebDriverWait wait = new WebDriverWait(driverMSD, TimeSpan.FromSeconds(60));

                        // Locate the element containing the problemDescription
                        IWebElement element = wait.Until(ExpectedConditions.ElementExists(By.CssSelector("textarea[aria-label='Problem Description']")));

                        // Extract the 'problemDescription' attribute
                        problemDesc = element.GetAttribute("title");

                        // Locate the input element using its aria-label
                        IWebElement mobilePhoneInput = wait.Until(ExpectedConditions.ElementExists(
                            By.CssSelector("input[aria-label='Mobile Phone']")
                        ));

                        // Retrieve the value of the title attribute
                        phoneNumber = mobilePhoneInput.GetAttribute("title");

                        IWebElement serialNumberInput = wait.Until(ExpectedConditions.ElementExists(
                            By.CssSelector("input[aria-label='Serial Number']")
                        ));

                        // Get the value of the title attribute
                        serialNumber = serialNumberInput.GetAttribute("title");

                        // Navigate to the URL
                        driverMSD.Navigate().GoToUrl("https://pcsupport.lenovo.com/us/en/warranty-lookup#/");
                        driverMSD.Manage().Window.FullScreen();

                        Thread.Sleep(1000);

                        // Wait for modal to close or any transition
                        if (IsButtonVisible(driverMSD, By.CssSelector(".button.blue-solid.modal-button-top40.btn_no")))
                        {
                            IWebElement modalButton = driverMSD.FindElement(By.CssSelector(".button.blue-solid.modal-button-top40.btn_no"));
                            var elements = driverMSD.FindElements(By.CssSelector(".button.blue-solid.modal-button-top40.btn_no"));

                            if (elements != null)
                            {
                                Thread.Sleep(1000);
                                modalButton.Click();
                            }
                        }

                        // Wait for modal to close or any transition
                        if (IsButtonVisible(driverMSD, By.CssSelector("span.close-cookie-remind")))
                        {
                            var waitCookies = new WebDriverWait(driverMSD, TimeSpan.FromSeconds(30));
                            var button1 = waitCookies.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("span.close-cookie-remind")));
                            button1.Click();

                            Thread.Sleep(800);
                        }

                        // Locate the input element using its class attribute
                        IWebElement inputField = driverMSD.FindElement(By.CssSelector("input.button-placeholder__input"));

                        Thread.Sleep(800);

                        inputField.SendKeys(serialNumber); // Send the text to the input field

                        Thread.Sleep(500);

                        // Locate the span element using its class attribute
                        IWebElement spanElement = driverMSD.FindElement(By.CssSelector("button.basic-search__suffix-btn.btn.btn-primary"));

                        // Click the span element
                        spanElement.Click();

                        Thread.Sleep(800);

                        // Now, adjust by scrolling up 300 pixels
                        ((IJavaScriptExecutor)driverMSD).ExecuteScript("window.scrollBy(0, 600);");

                        Thread.Sleep(800);

                        IJavaScriptExecutor js = (IJavaScriptExecutor)driverMSD;
                        IWebElement partsLink = driverMSD.FindElement(By.CssSelector("li.navtiles.mse-hover-menu a.vue-nav-menu[data-tab='psp-parts']"));
                        js.ExecuteScript("arguments[0].click();", partsLink);

                        Thread.Sleep(800);

                        driverMSD.Manage().Window.FullScreen();

                        // Now, adjust by scrolling up 300 pixels
                        ((IJavaScriptExecutor)driverMSD).ExecuteScript("window.scrollBy(0, 300);");
                        Thread.Sleep(1000);

                        if (domainUpDown1.Text == "Model")
                        {
                            //var wait1 = new WebDriverWait(driverMSD, TimeSpan.FromSeconds(20));
                            Thread.Sleep(800);

                            var locators = driverMSD.FindElements(By.CssSelector("div.el-tooltip.parts-tab"));
                            foreach (var el in locators)
                            {
                                var span = el.FindElement(By.CssSelector("span"));
                                if (span.Text == "Model")
                                {
                                    el.Click();
                                    break;
                                }
                            }
                            ((IJavaScriptExecutor)driverMSD).ExecuteScript("window.scrollBy(0, 300);");
                        }

                        // Wait until the element is present
                        WebDriverWait waitDownload = new WebDriverWait(driverMSD, TimeSpan.FromSeconds(10));

                        IWebElement iconDownload = driverMSD.FindElement(By.CssSelector("div.download-style .new-download-tip"));

                        // Scroll the page so that the element is in view
                        ((IJavaScriptExecutor)driverMSD).ExecuteScript("arguments[0].scrollIntoView(true);", iconDownload);

                        ((IJavaScriptExecutor)driverMSD).ExecuteScript("window.scrollBy(0, -150);");

                        IWebElement downloadElement = waitDownload.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.CssSelector("div.download-style .new-download-tip")));

                        // Click the element
                        downloadElement.Click();

                        // Wait for the download to complete
                        string expectedFileName = "Extracted Data.xlsx";

                        // Check if the file exists
                        if (File.Exists(expectedFileName))
                        {
                            // Delete the file
                            File.Delete(expectedFileName);
                        }

                        string downloadDirectory = AppDomain.CurrentDomain.BaseDirectory;


                        try
                        {
                            string downloadedFilePath = WaitForDownloadToComplete(downloadDirectory, TimeSpan.FromSeconds(10));// 30 seconds timeout
                            string renamedFilePath = Path.Combine(downloadDirectory, expectedFileName);// Rename the file
                            File.Move(downloadedFilePath, renamedFilePath);
                        }
                        catch (TimeoutException ex)
                        {
                            MessageBox.Show(ex.Message);
                        }

                        PopulateDropDownListFromExcel(expectedFileName);

                        PopulateCheckboxListFromExcel(expectedFileName);

                        comboBox1.Enabled = true;

                        driverMSD.Quit();

                        Process[] chromeDriverProcesses = Process.GetProcessesByName("chromedriver");
                        foreach (var chromeDriverProcess in chromeDriverProcesses)
                        {
                            chromeDriverProcess.Kill();
                        }

                        service.Dispose();

                        scraperButton.Enabled = true;

                        loadingScreen.Close();

                        foreach (var item in comboBox1.Items)
                        {
                            originalItems.Add(item.ToString());
                        }

                        MessageBox.Show("Data Exported Successfully");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error: {ex.Message}");
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Use Commercial for commercial projects
            //this.AutoScaleMode = AutoScaleMode.Dpi;
            this.AutoScaleMode = AutoScaleMode.Font;
            domainUpDown1.TextAlign = HorizontalAlignment.Center;
            domainUpDown1.Items.Add("As-Built");
            domainUpDown1.Items.Add("Model");
            loggedIn.Checked = true;
            domainUpDown1.SelectedIndex = 0;
            domainUpDown1.ReadOnly = true;
            comboBox1.Enabled = false;
            checkedListBox2.Enabled = false;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            comboBox2.SelectedIndex = 0;
            //pictureBox1.SendToBack();
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            string searchText = comboBox1.Text.ToLower();
            SearchCheckboxListUsingFilterFromExcel(excelFilePath, searchText);

        }

        private void notLoggedIn_CheckedChanged(object sender, EventArgs e)
        {
            //loggedIn.Checked = false;
        }

        private void loggedIn_CheckedChanged(object sender, EventArgs e)
        {
            //notLoggedIn.Checked = false;
        }

        private void domainUpDown1_SelectedItemChanged(object sender, EventArgs e)
        {
            //MessageBox.Show(domainUpDown1.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (checkedListBox2.Items.Count != 0)
            {
                // Get the directory where the application is running
                string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

                // Define a persistent directory for user data
                string folderName = @Path.Combine(baseDirectory, "Case Warehouse");

                string fileName = serialNumber + ".txt";

                string filePath = @Path.Combine(folderName, fileName);


                // Check if the folder exists, create it if it doesn't
                if (!Directory.Exists(folderName))
                {
                    Directory.CreateDirectory(folderName);
                }

                List<string> selectedItems = new List<string>();

                selectedItems.Clear();

                string result = "";

                foreach (object item in checkedListBox2.Items)
                {
                    if (item != null)
                    {
                        selectedItems.Add(item.ToString());
                    }
                }

                result = string.Join(" ", selectedItems);

                if (comboBox2.Text == "English")
                {
                    languageSelection = 0;
                }
                if (comboBox2.Text == "French")
                {
                    languageSelection = 1;
                }
                if (comboBox2.Text == "Spanish")
                {
                    languageSelection = 2;
                }

                string fileContent = "";
                //string fileContent = "XXXX SSR  INFO XXXX\r\n1. Textable Phone?> Y\r\n1.a. Textable number>" + phoneNumber + " \r\n1.b. Additional Cust Phone or Comments> NA\r\n2. Preferred Contact Method> Email\r\n3. Floor, Suite, or Apt number>NA \r\n4. Labor Only WO?> N\r\n5. Previous Repair (same symptom) in Last 30 days?> N\r\n6. ACTION PLAN> Install " + result + ".\r\n Please confirm no CID first.\r\nONLY USE NEEDED PARTS and return unused returnable parts to Lenovo.\r\nConfirm unit has the latest BIOS version.\r\n7. AP from PD GUIDE> XXXXX\r\nXXXXXXXX\r\n\r\n\r\n\r\n\r\nPrevious WO History: Y/N\r\nWarranty type: Onsite\r\nProblem Description:\r\n" + problemDesc + "\r\nActions Taken:\r\n• Work order type: Onsite\r\n• PD Code: XXXX\r\n• Parts Sent: " + result + "\r\n• Additional Information: NA\r\n• Emails Summary: N/Y\r\nSource:\r\n• PD Guide\r\n• Lenovo Website\r\n";
                if (languageSelection == 0)
                {
                    fileContent = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX\r\nIs this a Textable Phone Number? Y\r\nIf yes, textable phone number equals: " + phoneNumber + "\r\nCustomer Contact preference – Email, Text, or Phone:\r\nDoes Service Location Contain Suite Unit or Apartment - Y or N: \r\nand if yes, is that information contained within the contact card details? \r\nProblem Description Verified – Y or N?: Y\r\nIs this a Labor Only work order – Y or N? N\r\nACTION PLAN: <Install " + result + ">\r\nONLY USE NEEDED PARTS, and return any unused parts to Lenovo.\r\nLenovo Self Repair Guides: https://support.lenovo.com/us/en/solutions/HT516547 ß\r\nIBM Internal Lenovo PC Publisher Page: https://w3.ibm.com/w3publisher/us-ibm-lenovo-pc-support ß\r\nPlease be sure machine is using the latest BIOS version\r\nPrevious Repair (with the same symptom) within Last 30 days: Y or N? N\r\nAdditional Comments:\r\nXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX\r\n\r\n\r\n\r\n\r\nPrevious WO History: N\r\nWarranty type: Onsite\r\nProblem Description:\r\n" + problemDesc + "\r\nActions Taken:\r\n• Work order type: Onsite\r\n• PD Code: XXXX\r\n• Parts Sent: " + result + "\r\n• Additional Information: NA\r\n• Emails Summary: N/Y\r\nSource:\r\n• PD Guide\r\n• Lenovo Website\r\n";
                }
                if (languageSelection == 1)
                {
                    fileContent = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX\r\nEst-ce un numéro de téléphone pouvant être envoyé par SMS ? Y\r\n\r\nSi oui, le numéro de téléphone pouvant être envoyé par SMS est égal à : " + phoneNumber + "\r\nPréférence de contact client – ​​E-mail, SMS ou téléphone :\r\nL'emplacement du service contient-il une suite ou un appartement ? - O ou N : \r\net si oui, ces informations sont-elles contenues dans les détails de la carte de contact ? \r\nDescription du problème Vérifié – O ou N ? : O\r\nS'agit-il d'un ordre de travail uniquement pour la main-d'œuvre – O ou N ? N\r\nPLAN D'ACTION : <Installer " + result + ">\r\nUTILISEZ UNIQUEMENT LES PIÈCES NÉCESSAIRES et renvoyez toutes les pièces non utilisées à Lenovo.\r\nGuides d'auto-réparation Lenovo : https://support.lenovo.com/us/en/solutions/HT516547 ß\r\nPage de l'éditeur interne IBM Lenovo PC : https://w3.ibm.com/w3publisher/us-ibm-lenovo-pc-support ß\r\nAssurez-vous que la machine utilise la dernière version du BIOS\r\nRéparation antérieure (avec le même symptôme) au cours des 30 derniers jours : O ou N ? N\r\nCommentaires supplémentaires :\r\nXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX\r\n\r\n\r\n\r\n\r\nPrevious WO History: N\r\nWarranty type: Onsite\r\nProblem Description:\r\n" + problemDesc + "\r\nActions Taken:\r\n• Work order type: Onsite\r\n• PD Code: XXXX\r\n• Parts Sent: " + result + "\r\n• Additional Information: NA\r\n• Emails Summary: N/Y\r\nSource:\r\n• PD Guide\r\n• Lenovo Website\r\n";
                }
                if (languageSelection == 2)
                {
                    fileContent = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX\r\nEs este un número de teléfono al que se le pueden enviar mensajes de texto? Y\r\nEn caso afirmativo, el número de teléfono al que se puede enviar un mensaje de texto es igual a: " + phoneNumber + "\r\nPreferencia de contacto del cliente: correo electrónico, mensaje de texto o teléfono:\r\nLa ubicación del servicio contiene una unidad tipo suite o apartamento? - Sí o No: \r\nY si es así, ¿esa información está contenida en los datos de la tarjeta de contacto? \r\nDescripción del problema verificada: ¿Sí o no?: Sí\r\nEs esta una orden de trabajo solo para mano de obra? ¿S o N? N\r\nPLAN DE ACCIÓN: <Instalar " + result + ">\r\nUTILICE SOLO LAS PIEZAS NECESARIAS y devuelva las piezas no utilizadas a Lenovo.\r\nGuías de autorreparación de Lenovo: https://support.lenovo.com/us/en/solutions/HT516547 ß\r\nPágina interna del editor de PC Lenovo de IBM: https://w3.ibm.com/w3publisher/us-ibm-lenovo-pc-support ß\r\nAsegúrese de que la máquina esté usando la última versión del BIOS.\r\nReparación previa (con el mismo síntoma) en los últimos 30 días: ¿S o N? N\r\nComentarios adicionales:\r\nXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX\r\n\r\n\r\n\r\n\r\nPrevious WO History: N\r\nWarranty type: Onsite\r\nProblem Description:\r\n" + problemDesc + "\r\nActions Taken:\r\n• Work order type: Onsite\r\n• PD Code: XXXX\r\n• Parts Sent: " + result + "\r\n• Additional Information: NA\r\n• Emails Summary: N/Y\r\nSource:\r\n• PD Guide\r\n• Lenovo Website\r\n";
                }

                File.WriteAllText(filePath, fileContent);

                MessageBox.Show("SSR & Case Notes are Generated Succesfully!");

                // Open the file in the default text editor
                Process.Start(new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true // Ensures it uses the default editor
                });
            }
            else
            {
                MessageBox.Show("Please select parts first..");
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "Clear Filter")
            {
                PopulateCheckboxListFromExcel(excelFilePath);
            }
            else
            {
                PopulateCheckboxListUsingFilterFromExcel(excelFilePath);
            }
        }

        private void clearData_Click(object sender, EventArgs e)
        {
            checkedListBox2.Items.Clear();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void checkedListBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            int flag = 0;
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (checkedListBox1.GetItemCheckState(i) == CheckState.Checked)
                {
                    for (int j = 0; j < checkedListBox2.Items.Count; j++)
                    {
                        if (checkedListBox1.Items[i] == checkedListBox2.Items[j])
                        {
                            flag = 1;
                            break;
                        }
                    }
                    if (flag == 0)
                    {
                        checkedListBox2.Items.Add(checkedListBox1.Items[i]);
                        break;
                    }
                }
            }
        }

        private void checkedListBox2_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
