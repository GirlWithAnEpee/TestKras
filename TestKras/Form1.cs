using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace TestKras
{
    public partial class Form1 : Form
    {
        int parties, noms, mts; //количество партий, номенклатур и аппаратов, соответствено
        //для хранения информации о партиях, номенклатурах и аппаратах
        Dictionary<int, string> Parties = new Dictionary<int, string>();
        Dictionary<int, string> Noms = new Dictionary<int, string>();
        Dictionary<int, string> Mts = new Dictionary<int, string>();
        //для хранения информации о времени обработки
        List<Dictionary<int, int>> times = new List<Dictionary<int, int>>();
        public Form1()
        { 
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filename = Get_file();
            textBox1.Text = filename;
            Parties = Read_file(filename);
            parties = Parties.Count();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string filename = Get_file();
            textBox2.Text = filename;
            Noms = Read_file(filename);
            noms = Noms.Count();
            dataGridView1.RowCount = noms;
            int maxString = 0;
            for (int i = 0; i < noms; i++)

            {
                dataGridView1.Rows[i].HeaderCell.Value = Noms[i];
                if (Noms[i].Length > maxString)
                {
                    maxString = Noms[i].Length;
                }
            }
            dataGridView1.RowHeadersWidth = dataGridView1.RowHeadersWidth + (7 * maxString);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string filename = Get_file();
            textBox3.Text = filename;
            Mts = Read_file(filename);
            mts = Mts.Count();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string filename = Get_file();
            textBox4.Text = filename;
            times = Read_file_time(filename);
            for (int i = 0; i < noms; i++)
            {
                for (int j = 0; j < mts; j++)
                {
                    if (times[i].ContainsKey(j))
                    {
                        dataGridView1.Rows[i].Cells[j].Value = times[i][j];
                    }
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private string Get_file()
        {
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.InitialDirectory = Directory.GetCurrentDirectory();
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return "Error";
            // получаем выбранный файл
            else
                return openFileDialog1.FileName;
        }

        private Dictionary<int, string> Read_file(string filename)
        {
            Dictionary<int, string> result = new Dictionary<int, string>();
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
            Range firstColumn = ObjWorkSheet.UsedRange.Columns[1];
            System.Array myvalues1 = (System.Array)firstColumn.Cells.Value;
            Range secondColumn = ObjWorkSheet.UsedRange.Columns[2];
            System.Array myvalues2 = (System.Array)secondColumn.Cells.Value;
            string[] IDarray = myvalues1.OfType<object>().Select(o => o.ToString()).ToArray();
            string[] Namearray = myvalues2.OfType<object>().Select(o => o.ToString()).ToArray();
            // Переводим два массива в словарь - убрав первую строчку с названиями столбцов
            for (int i = 1; i < IDarray.Count(); i++)
            {
                result.Add(Convert.ToInt32(IDarray[i]), Namearray[i]);
            }
            //MessageBox.Show(result[11].ToString());
            return result;
        }
        private List<Dictionary<int, int>> Read_file_time(string filename)
        {
            List<Dictionary<int, int>> result = new List<Dictionary<int, int>>();
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
            Range firstColumn = ObjWorkSheet.UsedRange.Columns[1];
            System.Array myvalues1 = (System.Array)firstColumn.Cells.Value;
            Range secondColumn = ObjWorkSheet.UsedRange.Columns[2];
            System.Array myvalues2 = (System.Array)secondColumn.Cells.Value;
            Range thirdColumn = ObjWorkSheet.UsedRange.Columns[3];
            System.Array myvalues3 = (System.Array)thirdColumn.Cells.Value;
            string[] IDarray = myvalues1.OfType<object>().Select(o => o.ToString()).ToArray();
            string[] Namearray = myvalues2.OfType<object>().Select(o => o.ToString()).ToArray();
            string[] Timearray = myvalues3.OfType<object>().Select(o => o.ToString()).ToArray();
            //MessageBox.Show(IDarray.Count().ToString()+IDarray);
            // Переводим два массива в словарь - убрав первую строчку с названиями столбцов
            string tmp = IDarray[1];

            for (int i = 1; i < IDarray.Count(); i++)

            {
                Dictionary<int, int> tmplist = new Dictionary<int, int>();
                while (tmp == IDarray[i])
                {
                    tmplist.Add(Convert.ToInt32(Namearray[i]), Convert.ToInt32(Timearray[i]));
                    i++;
                    if (i == IDarray.Count())
                    {
                        break;
                    }
                }
                result.Add(tmplist);
                if (i == IDarray.Count())
                {
                    break;
                }
                tmp = IDarray[i];
                i--;
            }
            return result;
        }
    }
}
