using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template4439
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void NagumanovBtn_Click(object sender, RoutedEventArgs e)
        {
            _4439_Nagumanov window = new _4439_Nagumanov();
            window.Show();
        }

        private void importBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (spisok.xlsx)|*.xlsx",
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

            using(ISRPOEntities1 isrpoEntities = new ISRPOEntities1())
            {
                for (int i = 1; i < _rows; i++)
                {
                    isrpoEntities.var9.Add(new var9()
                    {
                        employe_id = list[i, 0],
                        post = list[i, 1],
                        fio = list[i, 2],
                        login = list[i, 3],
                        password = list[i, 4],
                        last_input = list[i, 5],
                        input_type = list[i, 6],
                    });
                    isrpoEntities.SaveChanges();
                }
            }
            MessageBox.Show("import success");
        }

        private void exportBtn_Click(object sender, RoutedEventArgs e)
        {
            List<var9> all_employe;

            using (ISRPOEntities1 isrpoEntities = new ISRPOEntities1())
            {
                all_employe = isrpoEntities.var9.ToList().OrderBy(s => s.input_type).ToList();
            }

            int sheetCount = all_employe.GroupBy(s => s.input_type).ToList().Count();

            var app = new Excel.Application();
            app.SheetsInNewWorkbook = sheetCount;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            var employeCategories = all_employe.GroupBy(s => s.input_type).ToList();

            for (int i = 0; i < sheetCount; i++)
            {
                string currentCategory = employeCategories[i].Key;

                int startRowIndex = 2;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = currentCategory;
                worksheet.Cells[1][startRowIndex] = "Код клiента";
                worksheet.Cells[2][startRowIndex] = "Должность";
                worksheet.Cells[3][startRowIndex] = "Логин";
                startRowIndex++;

                Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[3][1]];
                headerRange.Merge();
                headerRange.Value = currentCategory;
                headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                headerRange.Font.Italic = true;

                foreach (var9 employe in all_employe)
                {
                    var e1 = employe.input_type;
                    if (employe.input_type.Equals(currentCategory))
                    {
                        worksheet.Cells[1][startRowIndex] = employe.employe_id;
                        worksheet.Cells[2][startRowIndex] = employe.post;
                        worksheet.Cells[3][startRowIndex] = employe.login;
                        startRowIndex++;
                    }
                }
                Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1],
                worksheet.Cells[3][startRowIndex - 1]];
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                worksheet.Columns.AutoFit();
            }
            MessageBox.Show("export success");
            app.Visible = true;
        }
    }
}
