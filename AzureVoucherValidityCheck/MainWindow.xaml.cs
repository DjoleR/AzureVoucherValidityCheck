using System.Windows;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
//using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Remote;
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
                driver = new ChromeDriver(@"C:\Users\djord\Downloads\chromedriver_win32"); // Change this to the location where chrome driver is installed (extracted)                
                //driver = new EdgeDriver(@"C:\Users\djord\Downloads\edgedriver"); // Change this to the location where edge driver is installed (extracted)                
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
            voucherTextBox.Text = "";
        }

        private void CheckButton_Click(object sender, RoutedEventArgs e)
        {
            driver.Url = @"http://www.microsoftazurepass.com/";
            statusTextBlock.Text = "";
            statusTextBlock.Background = new SolidColorBrush(Colors.Transparent);
            var country = driver.FindElementById("ddlCountry");
            country.SendKeys("Switzerland");
            var voucher = driver.FindElementById("tbPromo");
            voucher.SendKeys(voucherTextBox.Text);
            var submit = driver.FindElementByClassName("btn");
            submit.Submit();

            bool error = CheckError(driver);
            if (error)
            {
                statusTextBlock.Text = "Voucher used!";
                statusTextBlock.Background = new SolidColorBrush(Colors.Red);
            }
            else
            {
                statusTextBlock.Text = "Voucher not used!";
                statusTextBlock.Background = new SolidColorBrush(Colors.LightGreen);
            }
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
    }
}
