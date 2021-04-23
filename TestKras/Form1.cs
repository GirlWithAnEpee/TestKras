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
        // список для подсчёта единиц номенклатур обработанных каждой печью
        List<List<int>> UnitstoMts = new List<List<int>>();
        //для хранения приоритетов обработки номенклатур в печах
        //Priority[id_nom]->[top_1_mts, top_2_mts, ...]
        List<List<int>> Priority = new List<List<int>>();
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
            button3.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string filename = Get_file();
            textBox3.Text = filename;
            Mts = Read_file(filename);
            mts = Mts.Count();
            dataGridView1.ColumnCount = mts;
            for (int i = 0; i < mts; i++)
            {
                dataGridView1.Columns[i].HeaderText = Mts[i];
            }
            button4.Enabled = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string filename = Get_file();
            textBox4.Text = filename;
            times = Read_file_time(filename);
            for (int i = 0; i < noms; i++)
            {
                for (int j = 0; j < mts; j++)
                // заполняем массив приоритетов
                // ставим большое время потому что неизвестно что в ячейке первой - может пустота
                // int minT == 1000;
                {
                    if (times[i].ContainsKey(j))
                    {
                        dataGridView1.Rows[i].Cells[j].Value = times[i][j];
                    }
                }
            }
            button1.Enabled = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private string Get_file()
        {
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.InitialDirectory = Directory.GetCurrentDirectory();
            openFileDialog1.FileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return "Error";
            // получаем выбранный файл
            else
                return openFileDialog1.FileName;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (Parties.Count() == 0)
                MessageBox.Show("Вы не загрузили данные о партиях!");
            
            else if (Noms.Count() == 0)
                MessageBox.Show("Вы не загрузили данные о номенклатурах!");
            else
            {
                // обработка партии
                // список для подсчёта времени работы каждой печи
                List<int> MtsTimes = new List<int>();

                // создаём вложенный исписок и заполняем оба нулями
                for (int i = 0; i < mts; i++)
                {
                    // тайминги
                    MtsTimes.Add(0);
                    // единицы номенклатур
                    List<int> tmpl = new List<int>();
                    for (int j = 0; j < noms; j++)
                    {
                        tmpl.Add(0);
                    }
                    UnitstoMts.Add(tmpl);
                }

                for (int i = 0; i < parties; i++)
                {
                    switch (Parties[i])
                    {
                        case "0":
                            // золото идёт в "золотую" печ
                            if (MtsTimes[1] < MtsTimes[0])
                            {
                                UnitstoMts[1][Convert.ToInt32(Parties[i])]++;
                                // заполняем словарь с партями инфой куда отправили партию
                                Parties[i] = Noms[Convert.ToInt32(Parties[i])] + "-> " + Mts[1] + " at " + MtsTimes[1];
                                // добавляем время работы печи
                                MtsTimes[1] += 20;

                            }
                            // иначе суём в серебрянную печ
                            else
                            {
                                UnitstoMts[0][Convert.ToInt32(Parties[i])]++;
                                // заполняем словарь с партями инфой куда отправили партию
                                Parties[i] = Noms[Convert.ToInt32(Parties[i])] + "-> " + Mts[0] + " at " + MtsTimes[0];
                                // добавляем время работы печи
                                MtsTimes[0] += 40;
                            }
                            break;
                        case "1":
                            //серебро сначала в печ 1 потом в печ 2 потом уже в 3
                            if (MtsTimes[0] < MtsTimes[1])
                            {
                                UnitstoMts[0][Convert.ToInt32(Parties[i])]++;
                                // заполняем словарь с партями инфой куда отправили партию
                                Parties[i] = Noms[Convert.ToInt32(Parties[i])] + "-> " + Mts[0] + " at " + MtsTimes[0];
                                // добавляем время работы печи
                                MtsTimes[0] += 20;
                            }
                            // иначе суём в серебрянную печ
                            else
                            {
                                if (MtsTimes[1] < MtsTimes[2])
                                {
                                    UnitstoMts[1][Convert.ToInt32(Parties[i])]++;
                                    // заполняем словарь с партями инфой куда отправили партию
                                    Parties[i] = Noms[Convert.ToInt32(Parties[i])] + "-> " + Mts[1] + " at " + MtsTimes[1];
                                    // добавляем время работы печи
                                    MtsTimes[1] += 30;
                                }
                                else
                                {
                                    UnitstoMts[2][Convert.ToInt32(Parties[i])]++;
                                    // заполняем словарь с партями инфой куда отправили партию
                                    Parties[i] = Noms[Convert.ToInt32(Parties[i])] + "-> " + Mts[2] + " at " + MtsTimes[2];
                                    // добавляем время работы печи
                                    MtsTimes[2] += 40;
                                }

                            }
                            break;
                        case "2":
                            if (MtsTimes[1] < MtsTimes[2])
                            {
                                UnitstoMts[1][Convert.ToInt32(Parties[i])]++;
                                // заполняем словарь с партями инфой куда отправили партию
                                Parties[i] = Noms[Convert.ToInt32(Parties[i])] + "-> " + Mts[1] + " at " + MtsTimes[1];
                                // добавляем время работы печи
                                MtsTimes[1] += 40;
                            }
                            else
                            {
                                UnitstoMts[2][Convert.ToInt32(Parties[i])]++;
                                // заполняем словарь с партями инфой куда отправили партию
                                Parties[i] = Noms[Convert.ToInt32(Parties[i])] + "-> " + Mts[2] + " at " + MtsTimes[2];
                                // добавляем время работы печи
                                MtsTimes[2] += 50;
                            }
                            break;

                        default:
                            break;
                    }
                }
  
                Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
                Workbook wb = oXL.Workbooks.Add(XlSheetType.xlWorksheet);
                Worksheet ws = (Worksheet)oXL.ActiveSheet;
                oXL.Visible = true;
                ws.Cells[1, 1] = "Партия";
                ws.Cells[1, 2] = "Действия";
                for (int i = 0; i < parties; i++)
                {
                    ws.Cells[i + 2, 1] = i.ToString();
                    ws.Cells[i + 2, 2] = Parties[i];
                }
                // разбивка по номенклатурам/печам

                // заголовки печей и металлов

                for (int i = 0; i < noms; i++)
                {
                    ws.Cells[parties + 4, 2 + i] = Mts[i];
                    ws.Cells[parties + 5 + i, 1] = Noms[i];
                }
                for (int i = 0; i < noms; i++)
                {
                    for (int j = 0; j < noms; j++)
                    {
                        ws.Cells[parties + 5 + j, 2 + i] = UnitstoMts[j][i];
                    }
                }
                wb.SaveAs("c:\\plan.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wb.Close();
            }
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
            ObjWorkBook.Close();
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
            ObjWorkBook.Close();
            return result;
        }
    }
}
