using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
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

            using(ISRPOEntities isrpoEntities = new ISRPOEntities())
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

            using (ISRPOEntities isrpoEntities = new ISRPOEntities())
            {
                all_employe = isrpoEntities.var9.ToList().OrderBy(s => s.input_type).ToList();
            }

            var app = new Excel.Application();
            app.SheetsInNewWorkbook = 2;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < 2; i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                //worksheet.Name = Convert.ToString(all_employe[i].input_type);
                worksheet.Cells[1][startRowIndex] = "Код клиента";
                worksheet.Cells[2][startRowIndex] = "Должность";
                worksheet.Cells[3][startRowIndex] = "Логин";
                startRowIndex++;

                var employeCategories = all_employe.GroupBy(s => s.input_type).ToList();

                foreach (var employe in employeCategories)
                {
                    if (employe.Key == all_employe[i].input_type)
                    {
                        Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][2], worksheet.Cells[3][2]];
                        headerRange.Merge();
                        headerRange.Value = all_employe[i].input_type;

                        startRowIndex++;

                        foreach (var9 var9 in all_employe)
                        {
                            if (employe.Key == var9.input_type)
                            {
                                worksheet.Cells[1][startRowIndex] = var9.employe_id;
                                worksheet.Cells[2][startRowIndex] = var9.post;
                                worksheet.Cells[3][startRowIndex] = var9.login;
                                startRowIndex++;
                            }
                        }
                    }
                    else
                    {
                        continue;
                    }
                }

            }
            MessageBox.Show("export success");
            app.Visible = true;
        }
    }
}
