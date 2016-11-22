using System.Windows;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
//using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Remote;
using System;
using Excel = Microsoft.Office.Interop.Excel;

using System.Runtime.InteropServices;
using System.Windows.Media;

namespace AzureVoucherValidityCheck
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private RemoteWebDriver driver = null;
        public MainWindow()
        {
            InitializeComponent();
            LoadKeys();
            try
            {
                driver = new ChromeDriver(@".");                
            }
            catch (DriverServiceNotFoundException)
            {
                MessageBox.Show("You need to install ChromeWebDriver in order to run this app!");
                statusTextBlock.Text = "You need to install ChromeWebDriver in order to run this app! Download it from https://chromedriver.storage.googleapis.com/index.html?path=2.25/";
                return;
            }

        }

        private void LoadKeys()
        {
           // voucherTextBox.Text = "";
        }

        private void CheckButton_Click(object sender, RoutedEventArgs e)
        {

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePathTextBox.Text);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

          
          

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i < xlRange.Rows.Count; i++)
            {

                driver.Url = @"http://www.microsoftazurepass.com/";
                Console.Write(xlRange.Cells[i, 1].Value2.ToString() + "\r\n");

                statusTextBlock.Text = "";
                //statusTextBlock.Background = new SolidColorBrush(Colors.Transparent);
               
                var country = driver.FindElementById("ddlCountry");
              
                country.SendKeys("Switzerland");
                var voucher = driver.FindElementById("tbPromo");
                voucher.SendKeys(xlRange.Cells[i, 1].Value2.ToString());
                var submit = driver.FindElementByClassName("btn");
                submit.Submit();

                bool error = CheckError(driver);
                if (error)
                {
                    statusTextBlock.Text = "Voucher used!";
                    statusTextBlock.Background = new SolidColorBrush(Colors.Red);
                    xlRange.Cells[i, 2] = "Used";
                }
                else
                {
                    statusTextBlock.Text = "Voucher not used!";
                    statusTextBlock.Background = new SolidColorBrush(Colors.LightGreen);
                    xlRange.Cells[i, 2] = "Valid";
                }
        
        
                //add useful things here!   

            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            /*
           
            */
        }

        private bool CheckError(RemoteWebDriver driver)
        {
            try
            {
                // If page has an element with id 'PromoCodeError' that means that the voucher is invalid (expired or invalid input)
                var errorCheck = driver.FindElementById("PromoCodeError");
                return true;
            }
            catch (NotFoundException)
            {
                return false;
            }
        }

        private void openVoucherFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();


            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel Files(*.xlsx)|*.xlsx|Excel Files(*.xls)|*.xls";
            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();
            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string filename = dlg.FileName;
                filePathTextBox.Text = filename;
            }
        }
    }
}
