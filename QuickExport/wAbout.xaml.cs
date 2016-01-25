using System.Windows;


namespace QuickExport
{
    public partial class wAbout : Window
    {
        public wAbout()
        {
            InitializeComponent();

            this.sName.Text = Ui.mName;
            this.sInfo.Text = "Version: " + Ui.mVer + "\n" +
                "Release date: " + Ui.mDate + "\n" +
                "Copyright © Sancho, 2016";
            this.sWeb.Text = Ui.mWebSite;
            this.sEmail.Text = "e-mail: " + Ui.mEmail;
        }

        private void cmClose_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            this.Close();
        }

        private void sWeb_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            System.Diagnostics.Process.Start(Ui.mWebSite);
            this.Close();
        }

    }
}
