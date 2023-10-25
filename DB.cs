using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using MySqlConnector;
using System.Windows;

namespace rsad
{
    internal class DB
    {
        //public string ConnectionString = "SERVER=37.140.192.78;charset = cp866;DATABASE=u1662611_rsad;UID=u1662611_admin;PASSWORD=vova_12345;connection timeout = 180;";
        public string ConnectionString = "SERVER = 37.140.192.78;" + "DATABASE= u1662611_raisad;" + "UID=u1662611_rrsad;" + "PASSWORD=rsad_062023;" + "connection timeout = 1800";
        public MySqlConnectionStringBuilder builder = new MySqlConnectionStringBuilder();
        public string how;
        public static DB db;
        public static DB GetDB() // получаем экземпляр бд
        {
            if (db == null)
            {
                db = new DB();
                return db;
            }
            else
            {
                return db;
            }
        }
        
        public void PostRequest(string path) //отправить запрос в бд
        {
            try
            {
                using (MySqlConnection connection = new MySqlConnection(ConnectionString))
                {
                    connection.Open();
                    MySqlCommand command = connection.CreateCommand();
                    command.CommandText = path;

                    if (command.ExecuteNonQuery() == 1) MessageBox.Show("Успешно!");
                    else MessageBox.Show("Произошла ошибка! Проверьте интерернет-соединение и повторите попытку снова");
                    connection.Close();
                }
            }
            catch
            {
                MessageBox.Show("Произошла ошибка! Проверьте подключение к интернету и повторите попытку");
            }
        }

        public int GetId(string path) //получение id объекта
        {
            try
            {
                using (MySqlConnection connection = new MySqlConnection(ConnectionString))
                {
                    String sql = path;

                    using (MySqlCommand command = new MySqlCommand(sql, connection))
                    {
                        connection.Open();
                        using (MySqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read() == true)
                            {
                                return reader.GetInt32(0);
                            }
                            else
                            {
                                connection.Close();
                                return 0;
                            }
                        }
                    }
                }
            }
            catch { return 0; }
        }
    }
}
