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

        private void BnTask_Click(object sender, RoutedEventArgs e)
        {

        }

        private void LogashinBtn_Click(object sender, RoutedEventArgs e)
        {
            _4439_Logashin window = new _4439_Logashin();
            window.Show();
            
        }

        private void _4439_Mingalimova_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
