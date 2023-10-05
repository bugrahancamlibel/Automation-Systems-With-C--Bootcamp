using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.IO;

namespace reminder_automation
{

    // This is a TODO list application that will remind you of your tasks. Saves data in excel file.
    
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string task = textBox1.Text;
            string date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            string excelFilePath = "Tasks.xlsx";

            using (var workbook = new XLWorkbook())
            {
                IXLWorksheet worksheet = null;

                // Check if the worksheet with the name "Tasks" already exists
                if (workbook.Worksheets.TryGetWorksheet("Tasks", out worksheet))
                {
                    // If it exists, use the existing worksheet
                    worksheet.Cell(1, 1).Value = "Task";
                    worksheet.Cell(1, 2).Value = "Date";
                }
                else
                {
                    // If it doesn't exist, create a new worksheet
                    worksheet = workbook.AddWorksheet("Tasks");
                    worksheet.Cell(1, 1).Value = "Task";
                    worksheet.Cell(1, 2).Value = "Date";
                }

                int row = worksheet.LastRowUsed()?.RowNumber() + 1 ?? 2;
                worksheet.Cell(row, 1).Value = task;
                worksheet.Cell(row, 2).Value = date;

                workbook.SaveAs(excelFilePath);
            }

            textBox1.Clear();
        }


        private void tabPage2_Enter_1(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();

            string excelFilePath = "Tasks.xlsx";

            if (File.Exists(excelFilePath))
            {
                using (var workbook = new XLWorkbook(excelFilePath))
                {
                    var worksheet = workbook.Worksheets.FirstOrDefault(w => w.Name == "Tasks");

                    if (worksheet != null)
                    {
                        foreach (var row in worksheet.RowsUsed().Skip(1))
                        {
                            string task = row.Cell(1).Value.ToString();
                            string date = row.Cell(2).Value.ToString();

                            dataGridView1.Rows.Add(task, date);
                        }
                    }
                }
            }
            else
            {
                // Handle the case where the Excel file doesn't exist.
                // You can display a message or take appropriate action.
                MessageBox.Show("Excel file not found. Please create the file or add tasks.");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                int selectedRowIndex = dataGridView1.SelectedRows[0].Index;
                string task = dataGridView1.Rows[selectedRowIndex].Cells["Task"].Value.ToString();

                // Remove the task from the DataGridView
                dataGridView1.Rows.RemoveAt(selectedRowIndex);

                // Remove the task from the Excel file
                using (var workbook = new XLWorkbook("Tasks.xlsx"))
                {
                    var worksheet = workbook.Worksheet("Tasks");

                    // Find and remove the row that matches the task
                    var rowToDelete = worksheet.RowsUsed()
                        .Where(row => row.Cell(1).Value.ToString() == task)
                        .FirstOrDefault();

                    if (rowToDelete != null)
                    {
                        rowToDelete.Delete();
                        workbook.SaveAs("Tasks.xlsx");
                    }
                }
            }
        }
    }
}
