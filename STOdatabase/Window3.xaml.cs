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
    /// Логика взаимодействия для Window3.xaml
    /// </summary>
    public partial class Window3 : Window
    {

        string dbFileName = "STO.accdb";
        public Window3()
        {
            InitializeComponent();
            LoadData();
            LoadComboBoxes();
        }
        public class UslugaItem
        {
            public int id { get; set; }
            public string Name { get; set; }
            public bool IsSelected { get; set; }
        }

        private void AddZakaz(object sender, RoutedEventArgs e)
        {
            Window1 window1 = new Window1();
            window1.Show();
            this.Close();
        }

        private void btnBack_Click(object sender, RoutedEventArgs e)
        {
            MainWindow mainwindow = new MainWindow();
            mainwindow.Show();
            this.Close();
        }

        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (OrdersGrid.SelectedItem == null)
            {
                MessageBox.Show("Пожалуйста, выберите заказ в таблице, который хотите изменить!", "Внимание");
                return;
            }
            DataRowView row = OrdersGrid.SelectedItem as DataRowView;
            int idZakaza = Convert.ToInt32(row["Номер заказа"]);
            if (date.SelectedDate == null || idAuto.SelectedIndex == -1 || idMaster.SelectedIndex == -1)
            {
                MessageBox.Show("Заполните все поля!", "Внимание");
                return;
            }

            string connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={AppDomain.CurrentDomain.BaseDirectory}{dbFileName};";
            try
            {
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    string q = "UPDATE Заказ SET [id Авто] = ?, [id Мастера] = ?, [Дата заказа] = ? WHERE [id Заказа] = ?";
                    OleDbCommand cmd = new OleDbCommand(q, conn);
                    cmd.Parameters.AddWithValue("?", ((DataRowView)idAuto.SelectedItem)["id Авто"]);
                    cmd.Parameters.AddWithValue("?", ((DataRowView)idMaster.SelectedItem)["id Мастера"]);
                    cmd.Parameters.AddWithValue("?", date.SelectedDate.Value);
                    cmd.Parameters.AddWithValue("?", idZakaza);
                    cmd.ExecuteNonQuery();

                    string deleteOldService = "DELETE FROM [Состав заказа] WHERE [id Заказа] = ?";
                    OleDbCommand cmdDel = new OleDbCommand(deleteOldService, conn);
                    cmdDel.Parameters.AddWithValue("?", idZakaza);
                    cmdDel.ExecuteNonQuery();

                    if (idUsluga.ItemsSource != null)
                    {


                        foreach (var item in idUsluga.Items)
                        {
                            if (item is UslugaItem usluga && usluga.IsSelected)
                            {
                                string insertService = "INSERT INTO [Состав заказа] ([id Заказа], [id Услуги])  VALUES(?, ?)";
                                OleDbCommand cmdIns = new OleDbCommand(insertService, conn);
                                cmdIns.Parameters.AddWithValue("?", idZakaza);
                                cmdIns.Parameters.AddWithValue("?", usluga.id);
                                cmdIns.ExecuteNonQuery();
                            }
                        }
                    }
                    MessageBox.Show("Заказ успешно обновлен!");
                    LoadData();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
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
                    З.[id Авто] AS CarID,
                    З.[id Мастера] AS MasterID,
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
                    З.[id Авто],
                    З.[id Мастера],
                    (М.[Фамилия] & ' ' & М.[Имя] & ' ' & М.[Отчество]),
                    А.[Номер машины],
                    З.[Дата заказа]
                        ORDER BY З.[Дата заказа] DESC";
                    OleDbDataAdapter adapter = new OleDbDataAdapter(queri, conn);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    OrdersGrid.ItemsSource = dt.DefaultView;
                    foreach(var column in OrdersGrid.Columns)
                    {
                        if(column.Header.ToString() == "CarID"||column.Header.ToString()=="MasterID")
                        {
                            column.Visibility = Visibility.Collapsed;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка подключения:\n" + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
        }

        private void OrdersGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (OrdersGrid.SelectedItem is DataRowView row)
            {
                date.SelectedDate = Convert.ToDateTime(row["Дата заказа"]);

                string car = row["Автомобиль"].ToString().Trim();
                foreach(DataRowView item in idAuto.Items)
                {
                    if (item["Номер машины"].ToString().Trim()==car)
                    {
                        idAuto.SelectedItem = item;
                        break;
                    }
                }

                string master = row["Мастер"].ToString().Trim();
                foreach (DataRowView item in idMaster.Items)
                {
                    if (item["ФИО"].ToString().Trim() == master)
                    {
                        idMaster.SelectedItem = item;
                        break;
                    }
                }

                int idZakaza = Convert.ToInt32(row["Номер заказа"]);
                LoadServicesForOrder(idZakaza);
            }
        }
        private void LoadServicesForOrder(int orderId)
        {
            string connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={AppDomain.CurrentDomain.BaseDirectory}{dbFileName};";
            List<int> selectedIds = new List<int>();
            try
            {
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    string q = "SELECT [id Услуги] FROM [Состав заказа] WHERE [id Заказа] = ?";
                    OleDbCommand cmd = new OleDbCommand(q, conn);
                    cmd.Parameters.AddWithValue("?", orderId);
                    OleDbDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        selectedIds.Add(Convert.ToInt32(reader["id Услуги"]));
                    }
                }
                if(idUsluga.ItemsSource !=null)
                {
                    foreach(var item in idUsluga.Items)
                    {
                        if(item is UslugaItem usluga)
                        {
                            usluga.IsSelected = selectedIds.Contains(usluga.id);
                        }
                    }
                }
                idUsluga.Items.Refresh();
            }
            catch(Exception ex)
            {
                MessageBox.Show("Ошибка загрузки услуг: " + ex.Message);
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
                    idAuto.DisplayMemberPath = "Номер машины";
                    idAuto.SelectedValuePath = "id Авто";

                    string qMaster = "SELECT [id Мастера], ([Фамилия] & ' ' & [Имя] & ' ' & [Отчество]) AS ФИО FROM Мастер";
                    OleDbDataAdapter daMaster = new OleDbDataAdapter(qMaster, conn);
                    DataTable dtMaster = new DataTable();
                    daMaster.Fill(dtMaster);
                    idMaster.ItemsSource = dtMaster.DefaultView;
                    idMaster.DisplayMemberPath = "ФИО";
                    idMaster.SelectedValuePath = "id Мастер";

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
                            IsSelected = false
                        });

                    }
                    idUsluga.ItemsSource = uslugiList;
                }
            }
            catch (Exception ex) { MessageBox.Show("Ошибка загрузки списков: " + ex.Message); }
        }

        private void OrdersGrid_AutoGeneratedColumns(object sender, EventArgs e)
        {
            foreach (var column in OrdersGrid.Columns)
            {
                if (column.Header.ToString() == "CarID" || column.Header.ToString() == "MasterID")
                {
                    column.Visibility = Visibility.Collapsed;
                }
            }
        }
    }
}
