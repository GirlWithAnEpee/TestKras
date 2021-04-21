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

namespace TestKras
{
    public partial class Form1 : Form
    {
        int parties, noms, mts; //количество партий, номенклатур и аппаратов, соответствено
        Dictionary<int, string> Parties, Noms, Mts; //для хранения информации о партиях, номенклатурах и аппаратах
        public Form1()
        {
            InitializeComponent();
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

        }
    }
}
