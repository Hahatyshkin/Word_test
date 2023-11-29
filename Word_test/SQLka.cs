using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;

namespace Word_test
{
    internal class SQLka
    {
        static SQLiteConnection connection;
        static SQLiteCommand command;
        static public bool Connect(string fileName)
        {
            try
            {
                connection = new SQLiteConnection("Data Source=" + fileName + ";Version=3; FailIfMissing=False");
                connection.Open();
                return true;
            }
            catch (SQLiteException ex)
            {
                Console.WriteLine($"Ошибка доступа к базе данных. Исключение: {ex.Message}");
                return false;
            }
        }
        static public void SQLk(string args)
        {
            if (Connect("LogBase.db"))
            {
                Console.WriteLine("Connected");
                command = new SQLiteCommand(connection)
                {
                    CommandText = "CREATE TABLE IF NOT EXISTS [Log]([msg]);"
                };
                command.ExecuteNonQuery();
                Console.WriteLine("Таблица создана");
                command.CommandText = "INSERT INTO Log (msg) VALUES (:msg)";
                command.Parameters.AddWithValue("name", args);
                command.ExecuteNonQuery();
            }
        }
    }
}
