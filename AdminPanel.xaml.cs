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
using System.Windows.Shapes;

namespace rsad
{
    /// <summary>
    /// Логика взаимодействия для AdminPanel.xaml
    /// </summary>
    public partial class AdminPanel : Window
    {

        public AdminPanel()
        {
            InitializeComponent();
            if(DataClass.myRole == "Сотрудник")
            {
                stafff.Visibility = Visibility.Hidden;
            }
            else
            {
                stafff.Visibility = Visibility.Visible;
            }
        }

        private void RequestBtn_Click(object sender, RoutedEventArgs e)
        {
            RequestWindow requestWindow = new RequestWindow();
            requestWindow.Show();
            this.Close();
        }

        private void ClientsBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ClientsWindow clients = new ClientsWindow();
                clients.Show();
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void StaffBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                StaffMainWindow mainWindow = new StaffMainWindow();
                mainWindow.Show();
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void BackBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MainWindow mainWindow = new MainWindow();
                mainWindow.Show();
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LkBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LkWindow mainWindow = new LkWindow();
                mainWindow.Show();
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
