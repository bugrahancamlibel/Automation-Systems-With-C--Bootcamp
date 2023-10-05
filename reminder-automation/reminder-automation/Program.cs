using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace reminder_automation
{
    internal static class Program
    {
        /// <summary>
        /// Uygulamanın ana girdi noktası.
        /// </summary>
        [STAThread]
        static void Main()
        {
            InitFiles();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        static void InitFiles()
        {
            // check if a directory is exist. if not, create it.
            if (!Directory.Exists("tasks"))
            {
                Directory.CreateDirectory("tasks");

                using (File.Create("tasks.xlsx")) { }
                // create a sheet called tasks
                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("tasks");
                worksheet.Cell("A1").Value = "Task";
                worksheet.Cell("B1").Value = "Date";
                workbook.SaveAs("tasks.xlsx");
            }
            else
            {
                // if directory exists, check if file exists. if not, create it.
                if (!File.Exists("tasks/tasks.xlsx"))
                {
                    using (File.Create("tasks.xlsx")) { }
                    // create a sheet called tasks
                    var workbook = new XLWorkbook();
                    var worksheet = workbook.Worksheets.Add("tasks");
                    worksheet.Cell("A1").Value = "Task";
                    worksheet.Cell("B1").Value = "Date";
                    workbook.SaveAs("tasks.xlsx");
                }
            }

            // check if a directory called sales is exist. if not, create it.
            if (!Directory.Exists("tasks"))
            {
                Directory.CreateDirectory("tasks");
                // create a file called sales.xlsx
                using (File.Create("tasks/tasks.xlsx")) { }
                // create a sheet called sales
                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("tasks");
                worksheet.Cell("A1").Value = "Task";
                worksheet.Cell("B1").Value = "Date";
                workbook.SaveAs("tasks/tasks.xlsx");
            }
            else
            {
                // if directory exists, check if file exists. if not, create it.
                if (!File.Exists("tasks/tasks.xlsx"))
                {
                    using (File.Create("tasks/tasks.xlsx")) { }
                    // create a sheet called tasks
                    var workbook = new XLWorkbook();
                    var worksheet = workbook.Worksheets.Add("tasks");
                    worksheet.Cell("A1").Value = "Task";
                    worksheet.Cell("B1").Value = "Date";
                    workbook.SaveAs("tasks/tasks.xlsx");

                }
            }
        }
    }
}
