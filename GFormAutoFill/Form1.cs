using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace GFormAutoFill
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Mở Chrome
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://docs.google.com/forms/d/e/1FAIpQLSduRG2Nlfb5AL5YTyolq7DaWv9kfg5ia911ywyjXlSVy52JIA/viewform");

            // Tạo WebDriverWait
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            // --- Nhập dữ liệu vào các ô input ---
            // Chờ và lấy tất cả input cùng class
            var inputs = wait.Until(d =>
            {
                var elements = d.FindElements(By.XPath("//input[contains(@class,'whsOnd zHQkBf')]"));
                return elements.Count > 0 ? elements : null;
            });

            if (inputs.Count >= 2)
            {
                inputs[0].SendKeys("your_name_here");
                inputs[1].SendKeys("6300630012");
            }

            // --- Chọn giá trị trong dropdown ---
            // Click vào dropdown để mở danh sách
            IWebElement dropdown = wait.Until(d =>
                d.FindElement(By.XPath("//div[@role='listbox']"))
            );
            dropdown.Click();

            // Chờ và chọn option "Thành phố Cần Thơ"
            IWebElement option = wait.Until(d =>
                d.FindElement(By.XPath("//div[contains(@class,'MocG8c') and @role='option']//span[text()='Thành phố Cần Thơ']"))
            );
            option.Click();

            // --- Click nút "Tiếp" ---
            IWebElement btnNext = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                By.XPath("//div[contains(@class, 'uArJ5e') and .//span[text()='Tiếp']]")
            ));
            btnNext.Click();
        }
    }
}
