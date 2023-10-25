using Microsoft.Office.Interop.Word;
using MySqlConnector;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web.Security;
using System.Web.UI.WebControls;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Xml.Linq;

namespace rsad
{
    /// <summary>
    /// Логика взаимодействия для RequestWindow.xaml
    /// </summary>
    public partial class RequestWindow : System.Windows.Window
    {
        MySqlConnection con = new MySqlConnection(DB.GetDB().ConnectionString);

        MySqlCommand cmd = new MySqlCommand();
        MySqlDataAdapter da = new MySqlDataAdapter();

        public int tgl = 0;
        List<string> toAddresses = new List<string>();
        

        public RequestWindow()
        {
            InitializeComponent();
            NonArch();
        }

        public void NonArch()
        {
            try
            {
                con.Open();
                string zapros = "SELECT Request.id AS 'ID', Request.type AS 'Тип заявки', Request.viezd AS 'Направление выезда', Request.addres AS 'Адрес', Request.count AS 'Кличество', Request.wish AS 'Пожелания', USERS.LastName AS 'Фамилия', USERS.Name AS 'Наме', USERS.SurName, Request.status AS 'Статус', Request.date AS 'Дата записи', Request.arch AS 'Архивные заявки' FROM Request INNER JOIN USERS ON Request.id_user = USERS.id_user WHERE Request.arch = 0";
                cmd = new MySqlCommand(zapros, con);
                da = new MySqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                System.Data.DataTable dt = new System.Data.DataTable();
                da.Fill(dt);
                ReqDGV.ItemsSource = dt.DefaultView;
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Arch()
        {
            try
            {
                con.Open();
                string zapros = "SELECT Request.id AS 'ID', Request.type AS 'Тип заявки', Request.viezd AS 'Направление выезда', Request.addres AS 'Адрес', Request.count AS 'Кличество', Request.wish AS 'Пожелания', USERS.LastName AS 'Фамилия', USERS.Name AS 'Наме', USERS.SurName, Request.status AS 'Статус', Request.date AS 'Дата записи', Request.arch AS 'Архивные заявки' FROM Request INNER JOIN USERS ON Request.id_user = USERS.id_user WHERE Request.arch = 1";
                cmd = new MySqlCommand(zapros, con);
                da = new MySqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                System.Data.DataTable dt = new System.Data.DataTable();
                da.Fill(dt);
                ReqDGV.ItemsSource = dt.DefaultView;
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void BtBack_Click(object sender, RoutedEventArgs e)
        {
            AdminPanel adminPanel = new AdminPanel();
            adminPanel.Show();
            this.Close();
        }

        private void EditReq_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //экселька
                // Создаем новый объект Excel
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = true;
                // Создаем новую книгу Excel
                Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
                // Получаем первый лист
                Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
                // Заполняем лист данными из DataGrid
                for (int i = 0; i < ReqDGV.Items.Count; i++)
                {
                    DataRowView rowView = (DataRowView)ReqDGV.Items[i];
                    DataRow row = rowView.Row;
                    for (int j = 0; j < row.ItemArray.Length; j++)
                    {
                        sheet.Cells[i + 1, j + 1] = row.ItemArray[j].ToString();
                    }
                }

                //Вызываем нашу созданную эксельку.
                excel.Visible = true;
                excel.UserControl = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            //try
            //{
            //    if (MessageBox.Show("Exel -Да, Word - Нет", "Подтверждение", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            //    {
            //        try
            //        {
            //            //экселька
            //            // Создаем новый объект Excel
            //            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            //            excel.Visible = true;
            //            // Создаем новую книгу Excel
            //            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
            //            // Получаем первый лист
            //            Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
            //            // Заполняем лист данными из DataGrid
            //            for (int i = 0; i < ReqDGV.Items.Count; i++)
            //            {
            //                DataRowView rowView = (DataRowView)ReqDGV.Items[i];
            //                DataRow row = rowView.Row;
            //                for (int j = 0; j < row.ItemArray.Length; j++)
            //                {
            //                    sheet.Cells[i + 1, j + 1] = row.ItemArray[j].ToString();
            //                }
            //            }

            //            //Вызываем нашу созданную эксельку.
            //            excel.Visible = true;
            //            excel.UserControl = true;
            //        }
            //        catch (Exception ex)
            //        {
            //            MessageBox.Show(ex.Message);
            //        }

            //    }
            //    else
            //    {
            //        try
            //        {
            //            // Создаем новый объект Word
            //            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            //            word.Visible = true;
            //            // Создаем новую документ Word
            //            Microsoft.Office.Interop.Word.Document document = word.Documents.Add();
            //            Microsoft.Office.Interop.Word.Table table = document.Tables.Add(document.Range(), ReqDGV.Items.Count + 1, ReqDGV.Columns.Count, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            //            Object behiavor = Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior;
            //            Object autoFitBehiavor = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed;
            //            // Заполняем таблицу данными из DataGrid
            //            for (int i = 0; i < ReqDGV.Columns.Count; i++)
            //            {
            //                table.Cell(1, i + 1).Range.Text = ReqDGV.Columns[i].Header.ToString();
                       
            //            }
            //            for (int i = 0; i < ReqDGV.Items.Count; i++)
            //            {
            //                DataRowView rowView = (DataRowView)ReqDGV.Items[i];
            //                DataRow row = rowView.Row;
            //                for (int j = 0; j < row.ItemArray.Length; j++)
            //                {
            //                    table.Cell(i + 2, j + 1).Range.Text = row.ItemArray[j].ToString();
            //                }
            //            }

            //            word.Activate();

            //            int dialogResult = word.Dialogs[Microsoft.Office.Interop.Word.WdWordDialog.wdDialogFilePrint].Show();

            //            word.ActiveDocument.Close();
            //        }
            //        catch (Exception ex)
            //        {
            //            MessageBox.Show(ex.Message);
            //        }
                    
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }
        //public void printing()
        //{
        //    try
        //    {
        //        try
        //        {
        //            using (var connection = new MySqlConnection(DB.GetDB().ConnectionString))
        //            {
        //                connection.Open();
        //                //List<string> toAddresses = new List<string>();

        //                var cmd = new MySqlCommand($"SELECT * FROM Request", connection);
        //                MySqlDataReader reader = cmd.ExecuteReader();

        //                while (reader.Read())
        //                {
        //                    while (reader.Read())
        //                    {
        //                        toAddresses.Add(reader.ToString());
        //                    }
        //                }
        //                reader.Close();

        //                connection.Close();
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(ex.Message);
        //        }


        //        try
        //        {
        //            DateTime dateTime = DateTime.Now.Date;
        //            string datenow = dateTime.ToString("dd.MM.yyyy");



        //            var helper = new WordHelper("kvitancia.docx");
        //            var items = new Dictionary<string, string>
        //        {
        //            {"<date_now>", datenow},
        //        };
        //            helper.Process(items);

        //            // helper.Process(items);

        //            MessageBox.Show("Квитанция составлена!");
        //        }
        //        catch
        //        {

        //        }


                
        //    }
        //    catch (Exception exc)
        //    {
        //        MessageBox.Show("Произошла ошибка! Подробности: " + exc.Message);
        //    }
        //}
        private void AddReq_Click(object sender, RoutedEventArgs e)
        {
            AddRequestWindow requestWindow = new AddRequestWindow();
            requestWindow.Show();
            this.Close();
        }

        private void AhchReq_Click(object sender, RoutedEventArgs e)
        {
            

            if( tgl == 0 )
            {
                Arch();
                tgl = 1;
            }
            else
            {
                NonArch();
                tgl=0;
            }
        }

        private void InAhchReq_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Переместить в архив?", "Подтверждение", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    DataClass.userID = (int)((DataRowView)ReqDGV.SelectedItems[0]).Row["ID"];

                    string query = $"UPDATE Request SET Request.arch = 1 WHERE  Request.id = '{DataClass.userID}' ";
                    DB.GetDB().PostRequest(query);
                    NonArch();
                }
                else
                {
                    NonArch();
                }
            }
            catch 
            {
                MessageBox.Show("Сначала необходимо указать строку!");
            }
        }
        private void SearchBtn_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void FilterReqBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //ViezdCMB.Text
                //     TypeCMB.Text;
                string viezd;
                string type;
                string and;
                string status;
                string aand;


                if (ViezdCMB.Text != "" )
                {
                    viezd = $"Request.viezd = '{ViezdCMB.Text} '";
                }
                else
                {
                    viezd = "";
                }

                if (TypeCMB.Text != "")
                {
                    type = $"Request.type = ' {TypeCMB.Text} '";
                }
                else
                {
                    type = "";
                }
                if (StatusCMB.Text != "")
                {
                    status = $"Request.status = ' {TypeCMB.Text} '";
                }
                else
                {
                    status = "";
                }

                if (viezd != "" && type != "") { and = "AND"; } else { and = ""; }


                if (viezd != "" && status != "") { aand = "AND"; } else { aand = ""; }

                try
                {
                    con.Open();
                    string zapros = $"SELECT Request.id AS 'ID', Request.type AS 'Тип заявки', Request.viezd AS 'Направление выезда', Request.addres AS 'Адрес', Request.count AS 'Кличество', Request.wish AS 'Пожелания', USERS.LastName AS 'Фамилия', USERS.Name AS 'Имя', USERS.SurName AS 'Отчество', Request.status AS 'Статус', Request.date AS 'Дата записи', Request.arch AS 'Архивные заявки' FROM Request INNER JOIN USERS ON Request.id_user = USERS.id_user WHERE {type} {and} {viezd} {aand} {status}";
                    cmd = new MySqlCommand(zapros, con);
                    da = new MySqlDataAdapter(cmd);
                    DataSet ds = new DataSet();
                    System.Data.DataTable dt = new System.Data.DataTable();
                    da.Fill(dt);
                    ReqDGV.ItemsSource = dt.DefaultView;
                    con.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    con.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void SbrosReqBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ViezdCMB.Text = "";
                TypeCMB.Text = "";
                StatusCMB.Text ="";
                NonArch();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
