using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace excel_interop
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            _xlApp = new Microsoft.Office.Interop.Excel.Application();            
        }
        private readonly Application _xlApp;
        private Workbook _xlBook = null;
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
                _xlBook?.Close();
                _xlApp.Quit();
            }
            base.Dispose(disposing);
        }

        private void checkBoxExcelVisible_CheckedChanged(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            BeginInvoke((MethodInvoker)delegate 
            {
                _xlApp.Visible = checkBoxExcelVisible.Checked; 
                checkBoxWorkbookOpen.Visible = _xlApp.Visible;
                Cursor = Cursors.Default;
            });
        }

        private void checkBoxWorkbookOpen_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxWorkbookOpen.Checked)
            {
                var filePath = System.IO.Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    "Excel",
                    "TestCells.xlsx");

                _xlBook = _xlApp.Workbooks.Open(filePath);
                Worksheet xlSheet = _xlBook.Sheets["Sheet1"];
                Range xlRange = xlSheet.Range["A1", "B4"];

                textBox1.Clear();
                for (int i = 1; i <= xlRange.Rows.Count; i++)
                {
                    var line = new List<string>();
                    for (int j = 1; j <= xlRange.Columns.Count; j++)
                    {
                        Range range = xlRange.Cells[i, j];
                        line.Add(range.Value2);
                    }
                    textBox1.AppendText(string.Join(" | ", line));
                    textBox1.AppendText(Environment.NewLine);
                }
            }
            else
            {
                _xlBook?.Close();
                _xlBook = null;
            }
        }
    }
}
