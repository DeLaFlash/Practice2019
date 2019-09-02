using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic.FileIO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Practice2
{
    public partial class Form1 : Form
    {
        /// <summary>
        /// Student mark collection
        /// </summary>
        public List<int> Marks = new List<int>();
        /// <summary>
        /// Marks amount field
        /// </summary>
        public int mrkCap = 0;
        /// <summary>
        /// "2" mark field   
        /// </summary>
        public int Cap2 = 0;
        /// <summary>
        /// "3" mark field
        /// </summary>
        public int Cap3 = 0;
        /// <summary>
        /// "4" mark field 
        /// </summary>
        public int Cap4 = 0;
        /// <summary>
        /// "5" mark field  
        /// </summary>
        public int Cap5 = 0;
        /// <summary>
        /// "2" mark percentage field
        /// </summary>
        public double perc2 = 0;
        /// <summary>
        /// "3" mark percentage field
        /// </summary>
        public double perc3 = 0;
        /// <summary>
        /// "4" mark percentage field
        /// </summary>
        public double perc4 = 0;
        /// <summary>
        /// "5" mark percentage field
        /// </summary>
        public double perc5 = 0;

        public Form1()
        {
            InitializeComponent();
        }
        /// <summary>
        /// "About" window opening button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button2_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.ShowDialog();
        }
        /// <summary>
        /// CSV file opening button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button1_Click(object sender, EventArgs e)
        {
            string rfname = @"C:\r.csv";
            OpenFileDialog open = new OpenFileDialog();
            open.InitialDirectory = "С:\\";
            open.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
            open.FilterIndex = 1;
            open.Title = "File opening";
            if (open.ShowDialog() == DialogResult.OK)
            {
                rfname = open.FileName;
                using (TextFieldParser fs = new TextFieldParser(rfname))
                {
                    fs.TextFieldType = FieldType.Delimited;
                    fs.SetDelimiters(",");
                    fs.ReadFields();
                    while (!fs.EndOfData)
                    {
                        string[] fields = fs.ReadFields();
                        if (fields[5].All(c => char.IsDigit(c)))
                            {
                            Marks.Add(Convert.ToInt32(fields[5]));
                            }
                        else Marks.Add(0);
                        
                    }
                }
            }
            //Бизнес-логика проекта
            for(int i=1;i<Marks.Count;i++)
            {
                if (Marks[i] != Marks[i - 1])
                    mrkCap++;
            }
            for (int i = 1; i < Marks.Count; i++)
            {
                if (Marks[i] != Marks[i - 1])
                {
                    if (Marks[i] >= 90 && Marks[i] <= 100)
                        Cap5++;
                    else if (Marks[i] >=75 && Marks[i] <= 89)
                        Cap4++;
                    else if (Marks[i] >= 60 && Marks[i] <= 74)
                        Cap3++;
                    else
                        Cap2++;
                }
                
            }
            checkBox1.Checked = true;
            perc5 = (Convert.ToDouble(Cap5) * 100) / Convert.ToDouble(mrkCap);
            label1.Text = "''5'' mark percentage: " + perc5.ToString();
            perc4 = (Convert.ToDouble(Cap4) * 100) / Convert.ToDouble(mrkCap);
            label2.Text = "''4'' mark percentage: " + perc4.ToString();
            perc3 = (Convert.ToDouble(Cap3) * 100) / Convert.ToDouble(mrkCap);
            label3.Text = "''3'' mark percentage" + perc3.ToString();
            perc2 = (Convert.ToDouble(Cap2) * 100) / Convert.ToDouble(mrkCap);
            label4.Text = "''2'' mark percentage: " + perc2.ToString();
        }
        /// <summary>
        /// Кнопка создания XLS и сохранения результатов
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button3_Click(object sender, EventArgs e)
        {
                    Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                    if (xlApp == null)
                    {
                        MessageBox.Show("Excel не установлен!");
                        return;
                    }
                    Excel.Workbook xlWorkBook;
                    Excel.Worksheet xlWorkSheet;
                    object misValue = System.Reflection.Missing.Value;
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    xlWorkSheet.Cells[1, 1] = "''5'' mark percentage";
                    xlWorkSheet.Cells[1, 2] = "''4'' mark percentage";
                    xlWorkSheet.Cells[1, 3] = "''3'' mark percentage";
                    xlWorkSheet.Cells[1, 4] = "''2'' mark percentage";
                    xlWorkSheet.Cells[2, 1] = perc5.ToString();
                    xlWorkSheet.Cells[2, 2] = perc4.ToString();
                    xlWorkSheet.Cells[2, 3] = perc3.ToString();
                    xlWorkSheet.Cells[2, 4] = perc2.ToString();
                    xlWorkBook.SaveAs("marks.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    MessageBox.Show("Файл marks.xls в папке проекта успешно создан!");
        }
    }
}
