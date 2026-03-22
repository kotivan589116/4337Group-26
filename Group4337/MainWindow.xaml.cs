using System.Windows;

namespace Group4337
{
    public partial class MainWindow : Window
    {
        public MainWindow()
            => InitializeComponent();

        private void Lunin_4337(object sender, RoutedEventArgs e)
        {
            var form = new _4337_Lunin();
            form.Show();
        }
    }
}