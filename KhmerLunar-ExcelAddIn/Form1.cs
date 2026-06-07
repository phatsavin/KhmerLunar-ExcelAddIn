using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace KhmerLunar_ExcelAddIn
{
    public partial class Form1 : Form
    {
        private MonthCalendar monthCalendar;
        private DateTime selectedDate;

        public Form1()
        {
            InitializeComponent();

            // Initialize calendar
            monthCalendar = new MonthCalendar();
            monthCalendar.MaxSelectionCount = 1;
            monthCalendar.DateSelected += MonthCalendar_DateSelected;

            // Position the calendar
            monthCalendar.Location = new System.Drawing.Point(20, 20);

            // Add calendar to the form
            this.Controls.Add(monthCalendar);
        }

        private void MonthCalendar_DateSelected(object sender, DateRangeEventArgs e)
        {
            selectedDate = e.Start;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (selectedDate != default(DateTime))
            {
                // Get the current Excel Application instance
                Application excelApp = Globals.ThisAddIn.Application;

                // Get the active workbook and worksheet
                Workbook activeWorkbook = excelApp.ActiveWorkbook;
                Worksheet activeWorksheet = (Worksheet)activeWorkbook.ActiveSheet;

                // Set the range (use ActiveCell to insert text in the selected cell)
                Range cell = activeWorksheet.Application.ActiveCell;

                // Insert the Khmer Lunar date into the selected cell
                cell.Value = KhmerLunar.getKhmerLunarString(selectedDate); // Ensure this method is implemented in your KhmerLunar class
                this.Close();
            }
            else
            {
                MessageBox.Show("Please select a date.");
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
