using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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

namespace DynamicOutlook
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SendEmailAsync();
        }

        /// <summary>
        /// Sends an email
        /// </summary>
        /// <returns>task</returns>
        private async Task SendEmailAsync()
        {
            string message;

            try
            {
                bool result = await OutlookHelper.SendAsync(Environment.UserName + "@local.com", "Test!", "test...");
                message = result ? "email sent successfully" : "failed to send email";
            }
            catch (COMException ex)
            {
                message = ex.GetBaseException().Message;
            }

            Dispatcher.Invoke(
                () =>
                {
                    MessageBox.Show(message);
                });
        }
    }
}
