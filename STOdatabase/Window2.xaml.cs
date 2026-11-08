using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
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

namespace STOdatabase
{
    /// <summary>
    /// Логика взаимодействия для Window2.xaml
    /// </summary>
    public partial class Window2 : Window
    {
        string dbFileName = "STO.accdb";
        public Window2()
        {
            InitializeComponent();
            LoadData();
        }

        private void btnBack_Click(object sender, RoutedEventArgs e)
        {
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
            this.Close();
        }

        private void LoadData()
        {
            string connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={AppDomain.CurrentDomain.BaseDirectory}{dbFileName};";
            try
            {
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    string queri = @"
                    SELECT 
                    З.[id Заказа] AS [Номер заказа],
                    (М.[Фамилия] & ' ' & М.[Имя] & ' ' & М.[Отчество]) AS [Мастер],
                    А.[Номер машины] AS [Автомобиль],
                    З.[Дата заказа],
                    SUM(У.[Стоимость услуги]) AS [Итоговая сумма]
                    FROM 
                    ((((Заказ AS З 
                        INNER JOIN Мастер AS М ON З.[id Мастера] = М.[id Мастера])
                        INNER JOIN Автомобиль AS А ON З.[id Авто] = А.[id Авто])
                        LEFT JOIN [Состав заказа] AS С ON З.[id Заказа] = С.[id Заказа])
                        LEFT JOIN Услуга AS У ON С.[id Услуги] = У.[id Услуги])  
                    GROUP BY
                    З.[id Заказа],
                    (М.[Фамилия] & ' ' & М.[Имя] & ' ' & М.[Отчество]),
                    А.[Номер машины],
                    З.[Дата заказа]
                        ORDER BY З.[Дата заказа] DESC";
                    OleDbDataAdapter adapter = new OleDbDataAdapter(queri, conn);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    OrdersGrid.ItemsSource = dt.DefaultView;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка подключения:\n" + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (OrdersGrid.SelectedItem == null)
            {
                MessageBox.Show("Пожалуйста, выберите заказ в таблице, который хотите удалить!", "Внимание");
                return;
            }
            DataRowView row = OrdersGrid.SelectedItem as DataRowView;
            int idZakaza = Convert.ToInt32(row["Номер заказа"]);
            string connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={AppDomain.CurrentDomain.BaseDirectory}{dbFileName};";
            try
            {
                MessageBoxResult result = MessageBox.Show($"Вы уверены, что хотите удалить заказ №{idZakaza}?", "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.No) return;
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    string delq1 = "DELETE FROM [Состав заказа] WHERE [id Заказа] = ?";
                    OleDbCommand cmdDel1 = new OleDbCommand(delq1, conn);
                    cmdDel1.Parameters.AddWithValue("?", idZakaza);
                    cmdDel1.ExecuteNonQuery();

                    string delq2 = "DELETE FROM Заказ WHERE [id Заказа] = ?";
                    OleDbCommand cmdDel2 = new OleDbCommand(delq2, conn);
                    cmdDel2.Parameters.AddWithValue("?", idZakaza);
                    cmdDel2.ExecuteNonQuery();

                    MessageBox.Show("Заказ успешно удален!", "Успех");
                    LoadData();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при удалении:\n" + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            Window3 window3 = new Window3();
            window3.Show();
            this.Close();
        }
    }
}
