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
        List<List<int>> times = new List<List<int>>();
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
            Dictionary<int, string> result= new Dictionary<int, string>();
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            int str = 2, col = 2, i=0;
            Microsoft.Office.Interop.Excel.Range cur = ObjWorkSheet.get_Range(str.ToString(), col.ToString());
            while (cur.ToString()!="")
            {
                i++;
                result.Add(Convert.ToInt32(cur), ObjWorkSheet.get_Range(2, 2).ToString());
                cur = ObjWorkSheet.get_Range(str.ToString() + i.ToString(), col.ToString() + i.ToString());
            }
            return result;
        }
    }
}
