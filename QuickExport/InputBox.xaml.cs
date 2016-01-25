using System.Windows;

namespace QuickExport
{
    public partial class InputBox : Window
    {
        public InputBox() {
            InitializeComponent();
            newName.Focus();
        }
        private void OK_Click(object sender, RoutedEventArgs e)
        {
            Ui.newName = newName.Text;
            this.Close();
        }
        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Ui.newName = "";
            this.Close();
        }
    }
}
