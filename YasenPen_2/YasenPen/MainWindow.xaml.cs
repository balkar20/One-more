using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data;
using System.Data.OleDb;
using System.Collections;
using System.IO;
using System.ComponentModel;
using Microsoft.Win32;

namespace YasenPen
{
    
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        bool isDoubleShow = false;

        DataTable dt = new DataTable();

        public List<string> itemsForVsp = new List<string>();
            
        public List<string> GetResponse()
        {
            return itemsForVsp;
        }

        public MainWindow()
        {
            InitializeComponent();
            

            this.Closing += MainWindow_Closing; 
            this.Loaded += MainWindow_Loaded;
            string fileText;
            txbx_vsp.Text = TypeEnroledRepository.GetInfo("vsp.txt");
            txbx_nsp.Text = TypeEnroledRepository.GetInfo("nsp.txt");
            txbx_dsp.Text = TypeEnroledRepository.GetInfo("dsp.txt");
            txbx_nfil.Text = TypeEnroledRepository.GetInfo("nfill.txt");
            txbx_ncbu.Text = TypeEnroledRepository.GetInfo("ncbu.txt");
            txbx_notd.Text = TypeEnroledRepository.GetInfo("notd.txt");
            txbx_acc.Text = TypeEnroledRepository.GetInfo("acc.txt");
            txbx_total_p.Text = TypeEnroledRepository.GetInfo("totalp.txt");
            txbx_total_s.Text = TypeEnroledRepository.GetInfo("totals.txt");
            txbx_nzp.Text = TypeEnroledRepository.GetInfo("nzp.txt");
            txbx_npp.Text = TypeEnroledRepository.GetInfo("npp.txt");
            txbx_dpp.Text = TypeEnroledRepository.GetInfo("dpp.txt");
            txbx_contract.Text = TypeEnroledRepository.GetInfo("contract.txt");
            //txbx_fisp.Text = TypeEnroledRepository.GetInfo("response.txt");




            StreamReader sr = new StreamReader(@"txtfile.txt");
            try
            {
                fileText = sr.ReadLine().ToString();
                txbx_way.Text = fileText;
            }
            catch
            {
                fileText = "";
                txbx_way.Text = fileText;
            }
            sr.Close();
        }

        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            ResponseRepository.WriteResponsivesInDoc(txbx_fisp.Text, "response.txt");
            TypeEnroledRepository.WriteInfoInDoc(txbx_vsp.Text, "vsp.txt");
            TypeEnroledRepository.WriteInfoInDoc(txbx_nsp.Text, "nsp.txt");
            TypeEnroledRepository.WriteInfoInDoc(txbx_dsp.Text, "dsp.txt");
            TypeEnroledRepository.WriteInfoInDoc(txbx_nfil.Text, "nfill.txt");
            TypeEnroledRepository.WriteInfoInDoc(txbx_ncbu.Text, "ncbu.txt");
            TypeEnroledRepository.WriteInfoInDoc(txbx_notd.Text, "notd.txt");
            TypeEnroledRepository.WriteInfoInDoc(txbx_acc.Text, "acc.txt");
            TypeEnroledRepository.WriteInfoInDoc(txbx_total_p.Text, "totalp.txt");
            TypeEnroledRepository.WriteInfoInDoc(txbx_total_s.Text, "totals.txt");
            TypeEnroledRepository.WriteInfoInDoc(txbx_nzp.Text, "nzp.txt");
            TypeEnroledRepository.WriteInfoInDoc(txbx_npp.Text, "npp.txt");
            TypeEnroledRepository.WriteInfoInDoc(txbx_dpp.Text, "dpp.txt");
            TypeEnroledRepository.WriteInfoInDoc(txbx_contract.Text, "contract.txt");
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            //ReadResponsievsInDoc(ref itemsForVsp, @"response.txt");
        }

        



        private void btn_go_Click(object sender, RoutedEventArgs e)
        {
            string header = "<HEADER>";
            string eod = "<EOD>";
            //Console.WriteLine("Введите номер списка:");
            string nsp;
            nsp = "<NSP>" + txbx_nsp.Text;
            //Console.WriteLine("Введите дату в формате 25.04.2017:");
            string dsp;
            dsp = "<DSP>" + txbx_dsp.Text;
            string nfil = "<NFIL>" + txbx_nfil.Text;
            string ncbu = "<NCBU>" + txbx_ncbu.Text;
            //Console.WriteLine("Введите номер структурного подразделения банка:");
            string notd;
            notd = "<NOTD>" + txbx_notd.Text;
            //Console.WriteLine("Введите номер счета плательщика с которого производятся перечисление денежных средств:");
            string acc;
            acc = "<ACC>" + txbx_acc.Text;
            string fisp = "<FISP>Балкаров А.В.";
            //Console.WriteLine("Введите общее количество получателей в списке:");
            string total_p;
            total_p = "<TOTAL_P>" + txbx_total_p.Text;
            //Console.WriteLine("Введите общюю сумму списка:");
            string total_s;
            total_s = "<TOTAL_S>" + txbx_total_s.Text;
            //Console.WriteLine("Введите назначение перечисления:");
            string nzp;
            nzp = "<NZP>" + txbx_nzp.Text;
            //Console.WriteLine("Введите номер платежного поручения:");
            string npp;
            npp = "<NPP>" + txbx_npp.Text;
            //Console.WriteLine("Введите дату платежного поручения на перечислене денежных средств:");
            string dpp;
            dpp = "<DPP>" + txbx_dpp.Text;
            //Console.WriteLine("Введите вид зачисляемого списка:");
            string vsp;
            vsp = "<VSP>" + txbx_vsp.Text;
            string contract = "<CONTRACT>" + txbx_contract.Text;
            string curency = "<CURRENCY>BYN";
            string delimiter = "<DELIMITER>";

            if (txbx_file.Text == "")
            {
                MessageBox.Show("Надо ввести путь к файлу в который мы запишем текст!");
                return;
            }
            if (dg_info.ItemsSource == null)
            {
                MessageBox.Show("В грид должна быть добавлена таблица - смотри инструкцию!");
                return;
            }

            string file = txbx_file.Text;

            

            List<string> items = new List<string>();

            foreach (DataRow row in dt.Rows)
            {
                // получаем все ячейки строки
                var cells = row.ItemArray;
                string strD;
                if (cells[3].ToString() != "" && cells[3].ToString() != " ")
                {
                    strD = String.Format("{0:0.00}", cells[3]);
                }
                else
                {
                    strD = "";
                }
                string newStrD = strD.Replace(",", ".");
                string str = cells[2].ToString() + " " + newStrD + " " + cells[1].ToString();
                if (strD != "")
                {
                    items.Add(str);
                }
            }

            try
            {
                string @fileWay = txbx_way.Text;
                StreamWriter sw = new StreamWriter(fileWay,true, Encoding.Default);
                sw.WriteLine(header);
                sw.WriteLine(nsp);
                sw.WriteLine(dsp);
                sw.WriteLine(nfil);
                sw.WriteLine(ncbu);
                sw.WriteLine(notd);
                sw.WriteLine(acc);
                sw.WriteLine(curency);
                sw.WriteLine(fisp);
                sw.WriteLine(total_p);
                sw.WriteLine(total_s);
                sw.WriteLine(nzp);
                sw.WriteLine(npp);
                sw.WriteLine(dpp);
                sw.WriteLine(vsp);
                sw.WriteLine(contract);

                sw.WriteLine(delimiter);

                for (int i = 0; i < items.Count; i++)
                {
                    sw.WriteLine(items[i]);
                }
                sw.WriteLine(eod);
                //Close the file
                sw.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Упс - что-то пошло не так........" + ex.Message);
            }
            finally
            {
                MessageBox.Show("Все данные были успешно записанны в текстовый файл!");
            }
            StreamWriter sw2 = new StreamWriter("txtfile.txt");

            sw2.WriteLine(txbx_way.Text);
            sw2.Close();
        }

        private void btn_findExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog opfd = new OpenFileDialog();
            if (opfd.ShowDialog() == true)
                txbx_file.Text = opfd.FileName.ToString();
        }

        private void btn_show_Click(object sender, RoutedEventArgs e)
        {
            string stringconn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txbx_file.Text + ";Extended Properties='Excel 12.0 Xml;HDR=YES;'";

            OleDbConnection conn = new OleDbConnection(stringconn);
            try
            {
                if (txbx_choice.Text != "")
                {
                    OleDbDataAdapter da = new OleDbDataAdapter("Select * from [" + txbx_choice.Text + "$]", conn);


                    da.Fill(dt);
                    dt.Columns[0].ColumnName = "№";
                    dt.Columns[1].ColumnName = "ФИО";
                    dt.Columns[2].ColumnName = "№СЧЕТА";
                    dt.Columns[3].ColumnName = "Сумма";


                    dg_info.ItemsSource = dt.DefaultView;
                    btn_go.IsEnabled = true;
                    
                }
                else
                    MessageBox.Show("Необходимо сначала выбрать файл Excel!!!");
            }
            catch
            {
                if (txbx_choice.Text == "")
                {
                    MessageBox.Show("Введите название листа Excel!!!");
                }
                else
                    MessageBox.Show("Файл Excel должен быть запущен!!!");
            }

            isDoubleShow = true;
        }
        
    }



    
}
