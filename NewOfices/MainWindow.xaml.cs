using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using System.Net;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Text.RegularExpressions;
using SalaryReport;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.RightsManagement;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.VisualStyles;
using System.Windows.Navigation;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using SalaryReport.Save;

namespace salary3Offices////////////////////////some
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private BackgroundWorker backgroundWorker;
        private BackgroundWorker backgroundWorker2;
        private BackgroundWorker backgroundWorker3;
        public static int port = 0;
        string sendSite;
        string pathToXml = Path.Combine(Directory.GetCurrentDirectory(), "data.xml");
        private string currencyUrl = @"http://www.nbrb.by/Services/XmlExRates.aspx?ondate=";
        public MainWindow()
        {
            InitializeComponent();
            backgroundWorker = ((BackgroundWorker) this.FindResource("backgroundWorker"));
            backgroundWorker2 = ((BackgroundWorker)this.FindResource("backgroundWorker2"));
            backgroundWorker3 = ((BackgroundWorker)this.FindResource("backgroundWorker3"));
            backgroundWorker.DoWork += BackgroundWorkerOnDoWork;
            backgroundWorker.RunWorkerCompleted += BackgroundWorkerOnRunWorkerCompleted;
            backgroundWorker2.DoWork += BackgroundWorkerOnDoWork;
            backgroundWorker3.DoWork += BackgroundWorkerOnDoWork;
            backgroundWorker2.RunWorkerCompleted += BackgroundWorkerOnRunWorkerCompleted;
            backgroundWorker3.RunWorkerCompleted += BackgroundWorkerOnRunWorkerCompleted;
            this.Closed += (sender, args) =>
            {
                SaveToXml(pathToXml);
            };
            RestoreFromXml(pathToXml);
        }

        private void BackgroundWorkerOnRunWorkerCompleted(object o, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                // Ошибка была сгенерирована обработчиком события DoWork
                MessageBox.Show(e.Error.Message, "Произошла ошибка");
            }
            else
            {
                SetCurencyInput result = (SetCurencyInput)e.Result;
                if (result.Input == "txbxDateZP")
                {
                    txbxCurrencyZP.Text = result.Currency;
                }
                else if (result.Input == "txbxDateHolliday")
                {
                    txbxCurrencyHolliday.Text = result.Currency;
                }
                else if (result.Input == "txbx_DateAvans")
                {
                    txbxCurrency.Text = result.Currency;
                }
            }
        }

        private void BackgroundWorkerOnDoWork(object o, DoWorkEventArgs doWorkEventArgs)
        {
            SetCurencyInput oldinput = (SetCurencyInput) doWorkEventArgs.Argument;

            //SetCurencyInput curency = new SetCurencyInput(oldinput.Input,oldinput.Date);
            oldinput.Currency = GetCurency(currencyUrl + oldinput.Currency);
            doWorkEventArgs.Result = oldinput;

            //DatePicker picker = ((DatePicker)sender);
            //string text = picker.Text;
            //string date = Convert.ToDateTime(text).ToString("MM/dd/yyyy");
            //StringBuilder bulder = new StringBuilder(date);
            //string result = bulder.Replace(".", "/").ToString();
            //Task<string> task = Task<string>.Factory.StartNew(() => GetCurency(currencyUrl + result));
            //if (picker.Name == "txbxDateZP")
            //{
            //    //txbxCurrencyZP.Text = ;
            //}
            //else if (picker.Name == "txbxDateHolliday")
            //{
            //    txbxCurrencyHolliday.Text = task.Result;
            //}
            //else if (picker.Name == "txbx_DateAvans")
            //{
            //    txbxCurrency.Text = task.Result;
            //}
        }

        private void CopyFilesInDirectory()
        {
            string pathFrom = Environment.CurrentDirectory + @"\..\..\Sent";

            string pathTo = txbxPathToCopy.Text;

            DirectoryInfo sourse = new DirectoryInfo(pathFrom);
            DirectoryInfo destin = new DirectoryInfo(pathTo + @"\");
            foreach (var item in sourse.GetFiles())
            {
                item.CopyTo(destin + item.Name, true);
            }
            foreach (var item in sourse.GetFiles())
            {
                item.Delete();
            }
        }

        private void BtnUpdate_Click(object sender, RoutedEventArgs e)
        {
            Helper.to.Clear();
            if (File.Exists(settingsFolder.Text))
            {
                Helper.ReadSettings(settingsFolder.Text);
                logs.Text = "Обновлено!! Настройки: от " + Helper.from + " подпись " + Helper.fromsign + "\n для " + Helper.to.Count + " сотрудников.";
            }
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = System.AppDomain.CurrentDomain.BaseDirectory;
            openFileDialog.Filter = "EXCEL Files (*.xls)|*.xls|EXCEL Files (*.xlsx)|*.xlsx";
            var result = openFileDialog.ShowDialog();
            if (result == false) return;
            fileFolder.Text = openFileDialog.FileName;
        }

        private void SendButton_Click(object sender, RoutedEventArgs e)
        {
            WarningWindow window = new WarningWindow();
            if (window.ShowDialog() == true)
            {
                if (window.DialogResult == true)
                {
                    return;
                }
            }

            string txtHolly = txbxDateHolliday.Text;

            if (port != 0)
            {
                Helper.login = new NetworkCredential(txbxLogin.Text, txbxPasssword.Text);
                Helper.port = port;
            }

            if (txbxDateZP.Text != "")
            {
                Helper.dateOfZpString = txbxDateZP.Text;
                Helper.currencyZP = txbxCurrencyZP.Text;

            }
            if (txbxDateHolliday.Text != "")
            {
                Helper.dateOfHollydayString = txbxDateHolliday.Text;
                Helper.curencyHolliday = txbxCurrencyHolliday.Text;
            }
            if (txbx_DateAvans.Text != "")
            {
                Helper.dateOfAvansString = txbx_DateAvans.Text;
                Helper.currency = txbxCurrency.Text;
            }

            if (String.IsNullOrEmpty(settingsFolder.Text))
            {
                MessageBox.Show("Выберите файл с настройками");
            }
            else if (String.IsNullOrEmpty(fileFolder.Text))
            {
                MessageBox.Show("Выберите файл с расчетными");
            }
            else
            {
                EnableDisableControls(false);
                try
                {
                    logs.Text = "Идет рассылка....";
                    logs.Text = File.ReadAllText(Helper.ConvertXslToCsv(settingsFolder.Text, fileFolder.Text, emailText.Text));
                }
                catch (Exception ex)
                {
                    logs.Text = ex.StackTrace;
                }
                EnableDisableControls(true);
            }
            CopyFilesInDirectory();
            //SaveToXml(pathToXml);
        }

        private void EnableDisableControls(bool isEnabled)
        {
            SettingsButton.IsEnabled = isEnabled;
            SendButton.IsEnabled = isEnabled;
            BrowseFile.IsEnabled = isEnabled;
            settingsFolder.IsEnabled = isEnabled;
            fileFolder.IsEnabled = isEnabled;
            ExitButton.IsEnabled = isEnabled;
        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            //btnUpdate.IsEnabled = true;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = System.AppDomain.CurrentDomain.BaseDirectory;
            openFileDialog.Filter = "Text Files (*.txt)|*.txt";
            var result = openFileDialog.ShowDialog();
            if (result == false) return;
            settingsFolder.Text = openFileDialog.FileName;
        }

        private void textBox1_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textbox = (TextBox)sender;

            Regex rgx = new Regex(Helper.patternDate);

            if (textbox.Text != "")
            {
                if (!rgx.IsMatch(textbox.Text))
                {
                    MessageBox.Show("Дата должна быть введена в формате : 25.03.2017 !!!");
                }
            }
        }

        private void textBoxDates_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textbox = (TextBox)sender;

            Regex rgx = new Regex(Helper.patternDate);

            if (textbox.Text != "")
            {
                if (!rgx.IsMatch(textbox.Text))
                {
                    MessageBox.Show("Дата должна быть введена в формате : 25.03.2017 !!!");
                }
            }

        }


        private void textBox1_TextChanged_1(object sender, TextChangedEventArgs e)
        {
            if (File.Exists(settingsFolder.Text))
            {
                Helper.ReadSettings(settingsFolder.Text);
                logs.Text = "Настройки: от " + Helper.from + " подпись " + Helper.fromsign + "\n для " + Helper.to.Count + " сотрудников.";
            }
        }

        private void button1_Click_1(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {

            RadioButton presed = (RadioButton)sender;

            //if(presed.Name != "rbtnAtezio")
            //{
            //    txbxLogin.IsEnabled = true;
            //    txbxPasssword.IsEnabled = true;
            //}
            //else
            //{
            //    txbxLogin.IsEnabled = false;
            //    txbxPasssword.IsEnabled = false;
            //}

            sendSite = presed.Content.ToString();
            if (sendSite == "gmail.com (smtp.gmail.com)")
            {
                port = 587;
            }
            else if (sendSite == "yandex.ru (smtp.yandex.ru)")
            {
                port = 587;
            }
            else if (sendSite == "mail.ru (smtp.mail.ru)")
            {
                port = 587;
            }
            else
            {
                port = 587;
            }

        }
        private void wayOfCopy_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
            var result = dialog.ShowDialog();

            txbxPathToCopy.Text = dialog.SelectedPath;
        }

        void SaveToXml(string fileName)
        {
            Datas datas = new Datas();
            datas.Currency = txbxCurrency.Text;
            datas.CurrencyHoliday = txbxCurrencyHolliday.Text;
            datas.CurrencyZP = txbxCurrencyZP.Text;
            datas.DateAvans = txbx_DateAvans.Text;
            datas.DateZp = txbxDateZP.Text;
            datas.DateHoliday = txbxDateHolliday.Text;
            datas.EmailText = emailText.Text;
            datas.FileFolder = fileFolder.Text;
            datas.Login = txbxLogin.Text;
            datas.Password = txbxPasssword.Text;
            datas.PathToCopy = txbxPathToCopy.Text;
            datas.SettingsFolder = settingsFolder.Text;

            XmlSerializer xs = new XmlSerializer(typeof(Datas));
            try
            {
                using (var fs = new FileStream(fileName, FileMode.Create))
                {
                    xs.Serialize(fs, datas);
                }
            }
            catch (Exception e)
            {
                string mes = e.Message;
                //Logger.Out("Не сериализовалось(");
            }

        }

        void RestoreFromXml(string fileName)
        {
            Datas datas = new Datas();
            XmlSerializer xs = new XmlSerializer(typeof(Datas));
            try
            {
                using (var fs = new FileStream(fileName, FileMode.Open))
                {
                    datas = (Datas)xs.Deserialize(fs);
                }
            }
            catch (Exception e)
            {
                //Logger.Out("Не десериализовалось(");
                return;
            }


            txbxCurrency.Text = datas.Currency;
            txbxCurrencyZP.Text = datas.CurrencyZP;
            txbxCurrencyHolliday.Text = datas.CurrencyHoliday;
            txbxDateHolliday.Text = datas.DateHoliday;
            txbxDateZP.Text = datas.DateZp;
            txbx_DateAvans.Text = datas.DateAvans;
            txbxLogin.Text = datas.Login;
            txbxPasssword.Text = datas.Password;
            txbxPathToCopy.Text = datas.PathToCopy;
            emailText.Text = datas.EmailText;
            settingsFolder.Text = datas.SettingsFolder;
            fileFolder.Text = datas.FileFolder;
        }

        string GetCurency(string url)
        {
            string cur = null;
            Servicer servicer = new Servicer();
            //Task<XmlDocument> task = new Task<XmlDocument>();
            XmlDocument doc = servicer.GetXmlCurencyData(url);
            XmlParser parser = new XmlParser();
            if (doc != null)
            {
                cur = parser.GetCurrency(doc);
            }

            return cur;
        }

        private  void SetCurrency(object sender, RoutedEventArgs e)
        {
            DatePicker picker = ((DatePicker)sender);
            string name = picker.Name;
            string text = picker.Text;
            string date = Convert.ToDateTime(text).ToString("MM/dd/yyyy");
            StringBuilder bulder = new StringBuilder(date);
            string result = bulder.Replace(".", "/").ToString();
            if (!backgroundWorker.IsBusy)
            {
                backgroundWorker.RunWorkerAsync(new SetCurencyInput(name, result));
            }
            else if (!backgroundWorker2.IsBusy)
            {
                backgroundWorker2.RunWorkerAsync(new SetCurencyInput(name, result));
            }
            else if(!backgroundWorker3.IsBusy)
            {
                backgroundWorker3.RunWorkerAsync(new SetCurencyInput(name, result));
            }

            
            
        }


        private void BtnClear_OnClick(object sender, RoutedEventArgs e)
        {
            Button but = (Button) sender;
            string btnName = but.Name;
            if (btnName == "btnClearAvans")
            {
                txbxCurrency.Text = "";
                txbx_DateAvans.Text = "";
            }
            else if (btnName == "btnClearZP")
            {
                txbxCurrencyZP.Text = "";
                txbxDateZP.Text = "";
            }
            else if (btnName == "btnClearHoliday")
            {
                txbxCurrencyHolliday.Text = "";
                txbxDateHolliday.Text = "";
            }
        }
    }
}
