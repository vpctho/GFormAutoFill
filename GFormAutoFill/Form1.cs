using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace GFormAutoFill
{
    public partial class Form1 : Form
    {
        string URL = "https://docs.google.com/forms/d/e/1FAIpQLSe1TuIpQiLMVg_JVTzPraOB2SYTKoUmXi6eQc3hxjNPcaprRw/viewform";
        int delay = 500;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Mở Chrome
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl(URL);

            // Tạo WebDriverWait
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            // --- Nhập dữ liệu vào các ô input ---
            // Chờ và lấy tất cả input cùng class
            var inputs = wait.Until(d =>
            {
                var elements = d.FindElements(By.XPath("//input[contains(@class,'whsOnd zHQkBf')]"));
                return elements.Count > 0 ? elements : null;
            });
            Thread.Sleep(delay); // Chờ thêm 1 chút để đảm bảo trang đã sẵn sàng
            if (inputs.Count >= 2)
            {
                inputs[0].SendKeys("your_name_here");
                inputs[1].SendKeys("6300630012");
            }

            Thread.Sleep(delay); // Chờ thêm 1 chút để đảm bảo trang đã sẵn sàng
            // Xã phường
            IWebElement xaphuong = wait.Until(d =>
                d.FindElement(By.XPath("//div[@role='listbox' and contains(@class, 'jgvuAb')]"))
            );
            xaphuong.Click();

            Thread.Sleep(delay);
            // Chờ và chọn option "xã Châu Thành"
            IWebElement xaChauThanh = wait.Until(d =>
                d.FindElement(By.XPath("//div[contains(@class,'MocG8c') and @role='option']//span[text()='Xã Châu Thành']"))
            );
            xaChauThanh.Click();

            // --- Click nút "Tiếp" ---
            //IWebElement btnNext = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
            //    By.XPath("//div[contains(@class, 'uArJ5e') and .//span[text()='Tiếp']]")
            //));
            //btnNext.Click();
        }
    }
}
