using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Group4337
{
    /// <summary>
    /// Логика взаимодействия для _4337_Lunin.xaml
    /// </summary>
    public partial class _4337_Lunin : Window
    {
        public _4337_Lunin()
        {
            InitializeComponent();
        }
     
        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };

            if (!(ofd.ShowDialog() == true))
                return;

            string[,] list;

            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;

            list = new string[_rows, _columns];

            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (ClientsEntities db = new ClientsEntities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    if (string.IsNullOrEmpty(list[i, 0])) continue;

                    DateTime birthDate = DateTime.Parse(list[i, 3]);

                    int age = DateTime.Today.Year - birthDate.Year;
                    if (birthDate.Date > DateTime.Today.AddYears(-age)) age--;

                    string category = "";
                    if (age >= 20 && age <= 29)
                        category = "Категория 1 (20-29)";
                    else if (age >= 30 && age <= 39)
                        category = "Категория 2 (30-39)";
                    else if (age >= 40)
                        category = "Категория 3 (40+)";
                    else
                        category = "Младше 20";

                    db.Clients.Add(new Clients()
                    {
                        Code = list[i, 0],
                        FullName = list[i, 1],
                        Email = list[i, 2],
                        BirthDate = birthDate,
                        Age = age,
                        AgeCategory = category
                    });
                }

                db.SaveChanges();
                MessageBox.Show($"Загружено записей: {db.Clients.Count()}");
            }
        }
        private void BnExport_Click(object sender, RoutedEventArgs e)
        {

            using (ClientsEntities db = new ClientsEntities())
            {
                var categories = db.Clients.GroupBy(c => c.AgeCategory).ToList();

                var app = new Excel.Application();
                app.SheetsInNewWorkbook = categories.Count;
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

                for (int i = 0; i < categories.Count; i++)
                {
                    var category = categories[i];
                    Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                    worksheet.Name = category.Key.Length > 31 ? category.Key.Substring(0, 31) : category.Key;

                    // ЗАГОЛОВКИ (как в методичке, но адаптировано)
                    worksheet.Cells[1, 1] = "Код клиента";
                    worksheet.Cells[1, 2] = "ФИО";
                    worksheet.Cells[1, 3] = "E-mail";
                    worksheet.Cells[1, 4] = "Дата рождения";
                    worksheet.Cells[1, 5] = "Возраст";

                    // Выделяем заголовки жирным (как в методичке)
                    Excel.Range headerRange = worksheet.Range[
                        worksheet.Cells[1, 1],
                        worksheet.Cells[1, 5]];
                    headerRange.Font.Bold = true;

                    // ДАННЫЕ
                    int row = 2;
                    foreach (var client in category)
                    {
                        worksheet.Cells[row, 1] = client.Code;
                        worksheet.Cells[row, 2] = client.FullName;
                        worksheet.Cells[row, 3] = client.Email;
                        worksheet.Cells[row, 4] = client.BirthDate?.ToShortDateString();
                        worksheet.Cells[row, 5] = client.Age;
                        row++;
                    }

                    // ФОРМУЛА ПОДСЧЕТА (как в методичке с СЧЁТ)
                    worksheet.Cells[row, 1].FormulaLocal = $"=СЧЁТ(E2:E{row - 1})";
                    worksheet.Cells[row, 1].Font.Bold = true;

                    // Оформление: границы и автоширина
                    Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[row, 5]];
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                    worksheet.Columns.AutoFit();
                }

                app.Visible = true;
            }
        }
    }
}
