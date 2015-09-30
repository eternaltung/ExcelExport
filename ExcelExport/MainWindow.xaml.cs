using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Drawing;
using System.Windows;
using ClosedXML.Excel;
using System.Data.OleDb;

namespace ExcelExport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Generate Random Student Data
        /// </summary>
        /// <param name="count">count</param>
        /// <returns>List<Student></returns>
        private List<Student> GenerateRandomData(int count)
        {
            Random rnd = new Random();
            List<Student> students = new List<Student>();
            for (int i = 0; i < count; i++)
            {
                students.Add(new Student()
                {
                    Name = "name" + i.ToString(),
                    Age = rnd.Next(10, 60),
                    Gender = rnd.NextDouble() >= 0.5,
                    ID = i,
                    phone = "0987654321",
                    Eng = rnd.Next(0, 100),
                    Math = rnd.Next(0, 100)
                });
            }
            return students;
        }

        private void NPOIButton_Click(object sender, RoutedEventArgs e)
        {
            List<Student> students = GenerateRandomData(100);
            
            XSSFWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("工作表_Test");

            //create the header row
            IRow headerRow = sheet.CreateRow(0);
            int count = 0;
            foreach (var item in typeof(Student).GetProperties())
            {
                headerRow.CreateCell(count++).SetCellValue(item.Name);
            }
            
            //set the value 
            for (int i = 1; i < students.Count() + 1; i++)
            {
                IRow row = sheet.CreateRow(i);
                count = 0;
                row.CreateCell(count++).SetCellValue(students[i - 1].ID);
                row.CreateCell(count++).SetCellValue(students[i - 1].Name);
                row.CreateCell(count++).SetCellValue(students[i - 1].Age);
                row.CreateCell(count++).SetCellValue(students[i - 1].Gender ? "男" : "女");
                row.CreateCell(count++).SetCellValue(students[i - 1].phone);
                row.CreateCell(count++).SetCellValue(students[i - 1].Eng);
                row.CreateCell(count).SetCellValue(students[i - 1].Math);
            }

            try
            {
                //save file
                SaveFileDialog saveDialog = new SaveFileDialog()
                {
                    DefaultExt = ".xlsx",
                    FileName = "Sample",
                    Filter = "Excel |*.xlsx"
                };
                if (saveDialog.ShowDialog() == true)
                {
                    using (FileStream MS = new FileStream(saveDialog.FileName,FileMode.Create))
                    {
                        workbook.Write(MS);
                    }
                }
            }
            catch (Exception ex)
            {
                StatusText.Text = ex.Message;
            }
        }

        private void EPPlusButton_Click(object sender, RoutedEventArgs e)
        {
            List<Student> students = GenerateRandomData(100);

            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet sheet = p.Workbook.Worksheets.Add("工作表_Test");

                //create the header row
                int count = 1;
                foreach (var item in typeof(Student).GetProperties())
                {
                    sheet.Cells[1, count].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[1, count].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    sheet.Cells[1, count++].Value = item.Name;
                }

                //set the value 
                for (int i = 2; i < students.Count() + 2; i++)
                {
                    count = 1;
                    sheet.Cells[i, count++].Value = students[i - 2].ID;
                    sheet.Cells[i, count++].Value = students[i - 2].Name;
                    sheet.Cells[i, count++].Value = students[i - 2].Age;
                    sheet.Cells[i, count++].Value = students[i - 2].Gender ? "男" : "女";
                    sheet.Cells[i, count++].Value = students[i - 2].phone;
                    sheet.Cells[i, count++].Value = students[i - 2].Eng;
                    sheet.Cells[i, count].Value = students[i - 2].Math;
                    sheet.Cells[i, count].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[i, count].Style.Fill.BackgroundColor.SetColor(students[i - 2].Math < 60 ? Color.Red : Color.Transparent);
                }
                sheet.Column(4).AutoFit();

                //condition format
                var ruleIcon = sheet.ConditionalFormatting.AddThreeIconSet(new ExcelAddress(2, 5, 10, 5), eExcelconditionalFormatting3IconsSetType.Signs);
                var colorformat = sheet.ConditionalFormatting.AddDatabar(new ExcelAddress(2, 6, 10, 6), Color.Blue);

                //add chart 
                ExcelWorksheet sheet2 = p.Workbook.Worksheets.Add("chart");
                ExcelChart chart = sheet2.Drawings.AddChart("NewChart", eChartType.Area3D);
                chart.Title.Text = "Title";
                chart.Series.Add(sheet.Cells[2, 6, 6, 6], sheet.Cells[2, 1, 6, 1]);
                chart.SetPosition(4, 1, 1, 0);
                chart.SetSize(800, 400);
                chart.View3D.DepthPercent = 60;
                chart.View3D.Perspective = 15;
                chart.View3D.RotX = 20;
                chart.View3D.RotY = 15;

                try
                {
                    //save file
                    SaveFileDialog saveDialog = new SaveFileDialog()
                    {
                        DefaultExt = ".xlsx",
                        FileName = "Sample",
                        Filter = "Excel |*.xlsx"
                    };
                    if (saveDialog.ShowDialog() == true)
                    {
                        File.WriteAllBytes(saveDialog.FileName, p.GetAsByteArray());
                        //p.SaveAs(new FileInfo(saveDialog.FileName));
                    }
                }
                catch (Exception ex)
                {
                    StatusText.Text = ex.Message;
                }
            }
        }

        private void ClosedXMLButton_Click(object sender, RoutedEventArgs e)
        {
            List<Student> students = GenerateRandomData(100);
            
            XLWorkbook workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("工作表_Test");
            int count = 1;
            foreach (var item in typeof(Student).GetProperties())
            {
                sheet.Cell(1, count).Style.Fill.PatternType = XLFillPatternValues.Solid;
                sheet.Cell(1, count).Style.Fill.BackgroundColor = XLColor.Yellow;
                sheet.Cell(1, count++).Value = item.Name;
            }

            for (int i = 2; i < students.Count() + 2; i++)
            {
                count = 1;
                sheet.Cell(i, count++).Value = students[i - 2].ID;
                sheet.Cell(i, count++).Value = students[i - 2].Name;
                sheet.Cell(i, count++).Value = students[i - 2].Age;
                sheet.Cell(i, count++).Value = students[i - 2].Gender ? "男" : "女";
                sheet.Cell(i, count++).Value = students[i - 2].phone;
                sheet.Cell(i, count++).Value = students[i - 2].Eng;
                sheet.Cell(i, count).Value = students[i - 2].Math;
                sheet.Cell(i, count).Style.Fill.PatternType = XLFillPatternValues.None;
                sheet.Cell(i, count).Style.Fill.BackgroundColor = students[i - 2].Math < 60 ? XLColor.Red : XLColor.Transparent;
            }
            sheet.Column(4).AdjustToContents();

            //condition format
            sheet.Range(2, 6, 10, 6).AddConditionalFormat().DataBar(XLColor.Orange).LowestValue().HighestValue();
            sheet.Range(2, 3, 10, 3).AddConditionalFormat().IconSet(XLIconSetStyle.ThreeTrafficLights2)
                .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, 45, XLCFContentType.Number)
                .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, 20, XLCFContentType.Number)
                .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, 10, XLCFContentType.Number);

            try
            {
                //save file
                SaveFileDialog saveDialog = new SaveFileDialog()
                {
                    DefaultExt = ".xlsx",
                    FileName = "Sample",
                    Filter = "Excel |*.xlsx"
                };
                if (saveDialog.ShowDialog() == true)
                {
                    workbook.SaveAs(saveDialog.FileName);
                }
            }
            catch (Exception ex)
            {
                StatusText.Text = ex.Message;
            }
        }

        private async void OleDBButton_Click(object sender, RoutedEventArgs e)
        {
            //此方式需要先安裝 http://www.microsoft.com/en-us/download/confirmation.aspx?id=23734
            List<Student> students = GenerateRandomData(100);

            try
            {
                SaveFileDialog saveDialog = new SaveFileDialog()
                {
                    DefaultExt = ".xlsx",
                    FileName = "Sample",
                    Filter = "Excel |*.xlsx"
                };
                if (saveDialog.ShowDialog() == true)
                {
                    string Connstring = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={saveDialog.FileName};Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                    using (OleDbConnection Conn = new OleDbConnection(Connstring))
                    {
                        await Conn.OpenAsync();

                        //create sheet
                        string sql = "CREATE TABLE Test (ID int,Name VarChar,Age int,Gender VarChar,phone VarChar,Eng int,Math int)";
                        using (OleDbCommand cmd = new OleDbCommand(sql, Conn))
                        {
                            await cmd.ExecuteNonQueryAsync();

                            //insert data
                            cmd.CommandText = "Insert into Test values (@ID,@Name,@Age,@Gender,@phone,@Eng,@Math)";
                            for (int i = 0; i < students.Count; i++)
                            {
                                cmd.Parameters.Clear();
                                cmd.Parameters.Add("@ID", OleDbType.Integer).Value = students[i].ID;
                                cmd.Parameters.Add("@Name", OleDbType.VarChar).Value = students[i].Name;
                                cmd.Parameters.Add("@Age", OleDbType.Integer).Value = students[i].Age;
                                cmd.Parameters.Add("@Gender", OleDbType.VarChar).Value = students[i].Gender ? "男" : "女";
                                cmd.Parameters.Add("@phone", OleDbType.VarChar).Value = students[i].phone;
                                cmd.Parameters.Add("@Eng", OleDbType.Integer).Value = students[i].Eng;
                                cmd.Parameters.Add("@Math", OleDbType.Integer).Value = students[i].Math;
                                await cmd.ExecuteNonQueryAsync();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                StatusText.Text = ex.Message;
            }
        }
    }
}
