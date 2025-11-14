using ExcelDataReader;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Data;

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
            // Kiểm tra xem file đã được chọn chưa
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("Chưa chọn file Excel!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return; // thoát khỏi hàm, không thực hiện phần xử lý tiếp theo
            }
            // Mở Chrome
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl(URL);

            // Tạo WebDriverWait
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

            // Đọc dữ liệu từ file Excel
            var data = ReadExcelToArray(@"D:\C#\mst.xlsx");
            foreach (var row in data)
            {
                //richTextBox1.AppendText(string.Join(" | ", row) + Environment.NewLine);
                // --- Nhập dữ liệu vào các ô input ---
                // Chờ và lấy tất cả input cùng class
                var inputs = wait.Until(d =>
                {
                    var elements = d.FindElements(By.XPath("//input[contains(@class,'whsOnd zHQkBf')]"));
                    return elements.Count > 0 ? elements : null;
                });
                Thread.Sleep(delay); // Chờ thêm 1 chút để đảm bảo trang đã sẵn sàng
                Thread.Sleep(delay); // Chờ thêm 1 chút để đảm bảo trang đã sẵn sàng
                if (inputs.Count >= 2)
                {
                    inputs[0].SendKeys(row[0]);
                    inputs[1].SendKeys(row[1]);
                }

                Thread.Sleep(delay); // Chờ thêm 1 chút để đảm bảo trang đã sẵn sàng
                                     // Xã phường
                IWebElement xaphuong = wait.Until(d =>
                    d.FindElement(By.XPath("//div[@role='listbox' and contains(@class, 'jgvuAb')]"))
                );
                xaphuong.Click();

                Thread.Sleep(delay);
                IWebElement xaChauThanh = wait.Until(d =>
                    d.FindElement(By.XPath($"//div[contains(@class,'MocG8c') and @role='option']//span[text()='{row[2]}']"))
                );
                xaChauThanh.Click();

                //Câu 2
                Thread.Sleep(delay);
                var option1 = wait.Until(d =>
                    d.FindElement(By.XPath($"//span[text()='{row[3]}']/ancestor::label"))
                );
                option1.Click();

                //Câu 3
                Thread.Sleep(delay);
                var option2 = wait.Until(d =>
                    d.FindElement(By.XPath($"//span[text()='{row[4]}']/ancestor::label"))
                );
                option2.Click();

                //Câu 4
                Thread.Sleep(delay);
                var option3 = wait.Until(d =>
                    d.FindElement(By.XPath($"//span[text()='{row[5]}']/ancestor::label"))
                );
                option3.Click();

                //Câu 5 - lựa chọn nhiều đáp án
                Thread.Sleep(delay);
                string cellValue = row[6]?.ToString();
                if (!string.IsNullOrEmpty(cellValue))
                {
                    // Tách các text bằng dấu ','
                    string[] texts = cellValue.Split(';');

                    foreach (string text in texts)
                    {
                        string trimmedText = text.Trim(); // loại bỏ khoảng trắng đầu/cuối

                        // Tìm option theo text và click
                        var option = wait.Until(d =>
                            d.FindElement(By.XPath($"//span[text()='{trimmedText}']/ancestor::label"))
                        );
                        option.Click();
                    }
                }




                //Câu 6
                Thread.Sleep(delay);
                var option5 = wait.Until(d =>
                    d.FindElement(By.XPath("//span[text()='Chưa biết cài đặt (cần hướng dẫn cài đặt)']/ancestor::label"))
                );
                option5.Click();

                //Câu 7
                Thread.Sleep(delay);
                var option6 = wait.Until(d =>
                    d.FindElement(By.XPath("//span[text()='Chưa có']/ancestor::label"))
                );
                option6.Click();

                //Câu 8
                Thread.Sleep(delay);
                var option7 = wait.Until(d =>
                    d.FindElement(By.XPath("//span[text()='Đã biết tới quy định mới']/ancestor::label"))
                );
                option7.Click();

                //Gửi
                Thread.Sleep(delay);
                var submit = wait.Until(d =>
                    d.FindElement(By.XPath("//span[text()='Gửi']/ancestor::div[contains(@class,'uArJ5e')]"))
                );
                submit.Click();

                //Hoàn tất
                Thread.Sleep(delay);
                Thread.Sleep(delay);
                var link = wait.Until(d =>
                    d.FindElement(By.XPath("//a[text()='Gửi ý kiến phản hồi khác']"))
                );
                Thread.Sleep(delay);
                Thread.Sleep(delay);
                link.Click();
            }


        }

        public static List<string[]> ReadExcelToArray(string filePath)
        {
            var resultList = new List<string[]>();

            // Bắt buộc cho ExcelDataReader
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true  // Dòng đầu là header, bắt đầu đọc từ dòng 2
                    }
                });

                DataTable table = dataSet.Tables[0];

                foreach (DataRow row in table.Rows)
                {
                    var arr = new string[10]; // A -> J

                    for (int i = 0; i < 10; i++)
                    {
                        arr[i] = row[i]?.ToString()?.Trim();
                    }

                    resultList.Add(arr);
                }
            }

            return resultList;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Chọn file Excel";
            openFileDialog1.Filter = "Excel Files|*.xlsx;*.xls";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string duongDanFile = openFileDialog1.FileName;
                textBox1.Text = duongDanFile; // hiển thị đường dẫn
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.StartPosition = FormStartPosition.CenterParent;
            form2.ShowDialog();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.TopMost = true;
        }
    }
}
