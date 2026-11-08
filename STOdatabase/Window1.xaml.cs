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
    /// Логика взаимодействия для Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        string dbFileName = "STO.accdb";
        public Window1()
        {
            InitializeComponent();
            LoadData();
            LoadComboBoxes();
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

        private void AddZakaz(object sender, RoutedEventArgs e)
        {
            if (date.SelectedDate == null || idAuto.SelectedValue == null || idMaster.SelectedValue == null )
            {
                MessageBox.Show("Заполните все поля!", "Внимание");
                return;
            }

            var uslugiList = idUsluga.ItemsSource as List<UslugaItem>;
            if(!uslugiList.Any(u=>u.IsSelected))
            {
                MessageBox.Show("Выберите хотя бы одну услугу!", "Внимание");
                return;
            }

            string connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={AppDomain.CurrentDomain.BaseDirectory}{dbFileName};";
            try
            {
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    int newOrderId = 0;

                    string insertQueri = @"
                    INSERT INTO Заказ ([id Авто], [id Мастера], [Дата заказа]) 
                    VALUES (?,?,?)";
                    OleDbCommand cmdInsert = new OleDbCommand(insertQueri, conn);
                    cmdInsert.Parameters.AddWithValue("?", idAuto.SelectedValue);
                    cmdInsert.Parameters.AddWithValue("?", idMaster.SelectedValue);
                    cmdInsert.Parameters.AddWithValue("?", date.SelectedDate.Value.ToShortDateString());
                    cmdInsert.ExecuteNonQuery();

                    string idQuery = "SELECT @@IDENTITY";
                    OleDbCommand cmdGetId = new OleDbCommand(idQuery, conn);
                    object result = cmdGetId.ExecuteScalar();
                    if (result != null && result != DBNull.Value)
                    {
                        newOrderId = Convert.ToInt32(result);
                    }

                    foreach (var usluga in uslugiList)
                    {
                        if(usluga.IsSelected)
                        {
                            string insertQueri2 = @"
                    INSERT INTO [Состав заказа] ([id Заказа], [id Услуги]) 
                    VALUES (?,?)";
                            OleDbCommand cmdInsert2 = new OleDbCommand(insertQueri2, conn);
                            cmdInsert2.Parameters.AddWithValue("?", newOrderId);
                            cmdInsert2.Parameters.AddWithValue("?", usluga.id);
                            cmdInsert2.ExecuteNonQuery();
                            
                        }
                    }
                    MessageBox.Show("Заказ успешно добавлен");

                    idAuto.SelectedIndex = -1;
                    idMaster.SelectedIndex = -1;
                    
                    date.SelectedDate = null;
                    foreach (var u in uslugiList) u.IsSelected = false;
                    idUsluga.Items.Refresh();

                    LoadData();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка подключения:\n" + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void LoadComboBoxes()
        {
            string connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={AppDomain.CurrentDomain.BaseDirectory}{dbFileName};";
            try
            {
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    string qAuto = "SELECT [id Авто], [Номер машины] FROM Автомобиль";
                    OleDbDataAdapter daAuto = new OleDbDataAdapter(qAuto, conn);
                    DataTable dtAuto = new DataTable();
                    daAuto.Fill(dtAuto);
                    idAuto.ItemsSource = dtAuto.DefaultView;

                    string qMaster = "SELECT [id Мастера], ([Фамилия] & ' ' & [Имя] & ' ' & [Отчество]) AS ФИО FROM Мастер";
                    OleDbDataAdapter daMaster = new OleDbDataAdapter(qMaster, conn);
                    DataTable dtMaster = new DataTable();
                    daMaster.Fill(dtMaster);
                    idMaster.ItemsSource = dtMaster.DefaultView;

                    string qUsluga = "SELECT [id Услуги], [Наименование услуги] FROM Услуга";
                    OleDbCommand cmdUsluga = new OleDbCommand(qUsluga, conn);
                    OleDbDataReader reader = cmdUsluga.ExecuteReader();
                    List<UslugaItem> uslugiList = new List<UslugaItem>();
                    while (reader.Read())
                    {
                        uslugiList.Add(new UslugaItem
                        {
                            id = Convert.ToInt32(reader["id Услуги"]),
                            Name = reader["Наименование услуги"].ToString(),
                            IsSelected =false
                        });
                        
                    }
                    idUsluga.ItemsSource = uslugiList;
                }
            }
            catch (Exception ex) { MessageBox.Show("Ошибка загрузки списков: " + ex.Message); }
        }

        private void btnBack_Click(object sender, RoutedEventArgs e)
        {
            MainWindow mainwindow = new MainWindow();
            mainwindow.Show();
            this.Close();
        }

        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            Window3 window3 = new Window3();
            window3.Show();
            this.Close();
        }
        public class UslugaItem
        {
            public int id { get; set; }
            public string Name { get; set; }
            public bool IsSelected { get; set; }
        }
    }
}
