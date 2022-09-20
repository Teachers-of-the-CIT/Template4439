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

        private void NagumanovBtn_Click(object sender, RoutedEventArgs e)
        {
            _4439_Nagumanov window = new _4439_Nagumanov();
            window.Show();
        }

        private void Minnullina_Click(object sender, RoutedEventArgs e)
        {
            new _4439_Minnullina().Show();
        }
    }
}
