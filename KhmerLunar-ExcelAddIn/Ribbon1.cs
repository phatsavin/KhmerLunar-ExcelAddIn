using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace KhmerLunar_ExcelAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            // Get the current Excel Application instance
            Application excelApp = Globals.ThisAddIn.Application;

            // Get the active workbook and active worksheet
            Workbook activeWorkbook = excelApp.ActiveWorkbook;
            Worksheet activeWorksheet = (Worksheet)activeWorkbook.ActiveSheet;

            // Set the range (use ActiveCell to insert the Khmer date)
            Range cell = activeWorksheet.Application.ActiveCell;

            // Insert the Khmer date into the active cell
            cell.Value = KhmerLunar.getKhmerLunarString(DateTime.Now); // Ensure this method is implemented in your KhmerLunar class
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 form = new Form1(); // Create an instance of Form2
            form.ShowDialog();

        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            // Get the current Excel Application instance
            Application excelApp = Globals.ThisAddIn.Application;

            // Get the active workbook and worksheet
            Workbook activeWorkbook = excelApp.ActiveWorkbook;
            Worksheet activeWorksheet = (Worksheet)activeWorkbook.ActiveSheet;

            // Set the range (use ActiveCell to insert text in the selected cell)
            Range cell = activeWorksheet.Application.ActiveCell;

            // Insert the Khmer full date into the selected cell
            cell.Value = KhmerLunar.GetKhmerFullDate(DateTime.Now); // Ensure this method is implemented in your KhmerLunar class
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Form2 form = new Form2(); // Create an instance of Form2
            form.ShowDialog();
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            Form3 form = new Form3(); // Create an instance of Form2
            form.ShowDialog();
            

        }
    }
}
