using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace It
{
    public partial class Form4 : Form
    {
        DataBase dataBase = new DataBase();
        string connectionString = "Data Source = ELIZZANKOVA\\SQLEXPRESS; Initial Catalog=Practice; Integrated Security=True";
       
        
        public Form4()
        {
            InitializeComponent();
            LoadMaterials();
        }

        private void LoadMaterials()
        {
            try
            {
                dataBase.openConnection();

                DateTime selectedDate = dateTimePicker.Value.Date;

                string query = @"SELECT S.Id_Товара, S.Наименование
                         FROM Sklad S
                         INNER JOIN Materials M ON S.Id_Материала = M.Id_Материала
                         WHERE S.Количество > 0 AND NOT EXISTS (
                               SELECT 1 FROM Income I 
                               WHERE I.Id_Товара = S.Id_Товара 
                               AND I.[Дата прихода] > @selectedDate)
                         ORDER BY S.Наименование";
                var command = new SqlCommand(query, dataBase.getConnection());
                command.Parameters.AddWithValue("@selectedDate", selectedDate);
                var reader = command.ExecuteReader();

                DataTable materialsTable = new DataTable();
                materialsTable.Load(reader);

                comboBox1.DataSource = materialsTable;
                comboBox1.DisplayMember = "Наименование";
                comboBox1.ValueMember = "Id_Товара";

                comboBox2.DataSource = materialsTable.Copy();
                comboBox2.DisplayMember = "Наименование";
                comboBox2.ValueMember = "Id_Товара";

                comboBox3.DataSource = materialsTable.Copy();
                comboBox3.DisplayMember = "Наименование";
                comboBox3.ValueMember = "Id_Товара";

                comboBox4.DataSource = materialsTable.Copy();
                comboBox4.DisplayMember = "Наименование";
                comboBox4.ValueMember = "Id_Товара";

                comboBox1.SelectedIndex = -1;
                comboBox2.SelectedIndex = -1;
                comboBox3.SelectedIndex = -1;
                comboBox4.SelectedIndex = -1;

                reader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке материалов: " + ex.Message);
            }
            finally
            {
                dataBase.closeConnection();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selectedMaterialId = Convert.ToInt32(comboBox1.SelectedValue);
            UpdateMaterialInfo(selectedMaterialId, txtprice1, txtqua1);
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selectedMaterialId = Convert.ToInt32(comboBox2.SelectedValue);
            UpdateMaterialInfo(selectedMaterialId, txtprice2, txtqua2);
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selectedMaterialId = Convert.ToInt32(comboBox3.SelectedValue);
            UpdateMaterialInfo(selectedMaterialId, txtprice3, txtqua3);
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selectedMaterialId = Convert.ToInt32(comboBox4.SelectedValue);
            UpdateMaterialInfo(selectedMaterialId, txtprice4, txtqua4);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int selectedMaterialId1 = Convert.ToInt32(comboBox1.SelectedValue);
            UpdateMaterialInfo(selectedMaterialId1, txtprice1, txtqua1);

            int selectedMaterialId2 = Convert.ToInt32(comboBox2.SelectedValue);
            UpdateMaterialInfo(selectedMaterialId2, txtprice2, txtqua2);

            int selectedMaterialId3 = Convert.ToInt32(comboBox3.SelectedValue);
            UpdateMaterialInfo(selectedMaterialId3, txtprice3, txtqua3);

            int selectedMaterialId4 = Convert.ToInt32(comboBox4.SelectedValue);
            UpdateMaterialInfo(selectedMaterialId4, txtprice4, txtqua4);
        }

        private void UpdateMaterialInfo(int materialId, TextBox priceTextBox, TextBox quantityTextBox)
        {
            try
            {
                dataBase.openConnection();

                string query = @"SELECT S.Цена, S.Количество
                 FROM Sklad S
                 INNER JOIN Materials M ON S.Id_Материала = M.Id_Материала
                 WHERE S.Id_Товара = @materialId
                 AND NOT EXISTS (
                       SELECT 1 FROM Income I 
                       WHERE I.Id_Товара = S.Id_Товара 
                       AND I.[Дата прихода] > @selectedDate)";
                var command = new SqlCommand(query, dataBase.getConnection());
                command.Parameters.AddWithValue("@materialId", materialId);
                command.Parameters.AddWithValue("@selectedDate", dateTimePicker.Value.Date);

                var reader = command.ExecuteReader();

                if (reader.Read())
                {
                    int price = Convert.ToInt32(reader["Цена"]);
                    int quantity = Convert.ToInt32(reader["Количество"]);

                    priceTextBox.Text = price.ToString();
                    quantityTextBox.Text = quantity.ToString();
                }
                else
                {
                    priceTextBox.Text = string.Empty;
                    quantityTextBox.Text = string.Empty;
                }

                reader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при обновлении информации о материале: " + ex.Message);
            }
            finally
            {
                dataBase.closeConnection();
            }
        }



        private void Add_Click(object sender, EventArgs e)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    int documentNumber = int.Parse(txtDocumentNumber.Text);
                    DateTime documentDate = dateTimePicker.Value.Date;

                    int totalCost = 0;

                    for (int i = 1; i <= 4; i++)
                    {
                        ComboBox comboBox = Controls.Find("comboBox" + i, true).FirstOrDefault() as ComboBox;
                        TextBox quantityTextBox = Controls.Find("txtqua" + i, true).FirstOrDefault() as TextBox;
                        TextBox priceTextBox = Controls.Find("txtprice" + i, true).FirstOrDefault() as TextBox;

                        int materialId = Convert.ToInt32(comboBox.SelectedValue);
                        int quantity = int.Parse(quantityTextBox.Text);
                        int price = int.Parse(priceTextBox.Text);

                        if (!IsMaterialAvailable(connection, materialId, quantity, price))
                        {
                            MessageBox.Show("Недостаточное количество материала на складе.");
                            return;
                        }

                        int cost = quantity * price;
                        totalCost += cost;

                        string insertQuery = $"INSERT INTO Outcome (Id_Товара, Id_Материала, Наименование, Количество, Цена, [Дата расхода], [Номер документа расхода])" +
                            $"VALUES (@materialId, (SELECT Id_Материала FROM Sklad WHERE Id_Товара = @materialId), (SELECT Наименование FROM Sklad WHERE Id_Товара = @materialId), @quantity, @price, @documentDate, @documentNumber)";
                        SqlCommand insertCommand = new SqlCommand(insertQuery, connection);
                        insertCommand.Parameters.AddWithValue("@materialId", materialId);
                        insertCommand.Parameters.AddWithValue("@quantity", quantity);
                        insertCommand.Parameters.AddWithValue("@price", price);
                        insertCommand.Parameters.AddWithValue("@documentDate", documentDate);
                        insertCommand.Parameters.AddWithValue("@documentNumber", documentNumber);
                        insertCommand.ExecuteNonQuery();

                        string updateQuery = $"UPDATE Sklad SET Количество = Количество - @quantity WHERE Id_Товара = @materialId";
                        SqlCommand updateCommand = new SqlCommand(updateQuery, connection);
                        updateCommand.Parameters.AddWithValue("@quantity", quantity);
                        updateCommand.Parameters.AddWithValue("@materialId", materialId);
                        updateCommand.ExecuteNonQuery();
                    }

                    txtTotalCost.Text = totalCost.ToString();

                    MessageBox.Show("Данные успешно сохранены.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при сохранении данных: " + ex.Message);
            }
        }

        private bool IsMaterialAvailable(SqlConnection connection, int materialId, int quantity, int price)
        {
            string query = "SELECT Количество FROM Sklad WHERE Id_Товара = @materialId";
            SqlCommand command = new SqlCommand(query, connection);
            command.Parameters.AddWithValue("@materialId", materialId);
            int availableQuantity = (int)command.ExecuteScalar();
            return (availableQuantity >= quantity);
        }





        private void Clear_Click(object sender, EventArgs e)
        {
            txtprice1.Text = string.Empty;
            txtprice2.Text = string.Empty;
            txtprice3.Text = string.Empty;
            txtprice4.Text = string.Empty;

            txtqua1.Text = string.Empty;
            txtqua2.Text = string.Empty;
            txtqua3.Text = string.Empty;
            txtqua4.Text = string.Empty;
            txtDocumentNumber.Text= string.Empty;

            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;

            txtTotalCost.Text = "";
        }

        private void word_Click(object sender, EventArgs e)
        {
            try
            {
                string fileName = "Расход.docx";
                using (WordprocessingDocument document = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = document.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());

                    // Добавление заголовка
                    Paragraph titleParagraph = body.AppendChild(new Paragraph());
                    Run titleRun = titleParagraph.AppendChild(new Run());
                    titleRun.AppendChild(new Text("Номер документа: " + txtDocumentNumber.Text));
                    titleRun.AppendChild(new Break());
                    titleRun.AppendChild(new Text("Дата документа: " + dateTimePicker.Value.ToShortDateString()));
                    titleRun.AppendChild(new Break());
                    titleRun.AppendChild(new Break());

                    // Добавление списка товаров
                    Paragraph listParagraph = body.AppendChild(new Paragraph());
                    Run listRun = listParagraph.AppendChild(new Run());
                    listRun.AppendChild(new Text("Список товаров:"));
                    listRun.AppendChild(new Break());
                    listRun.AppendChild(new Break());

                    for (int i = 1; i <= 4; i++)
                    {
                        ComboBox comboBox = Controls.Find("comboBox" + i, true).FirstOrDefault() as ComboBox;
                        TextBox quantityTextBox = Controls.Find("txtqua" + i, true).FirstOrDefault() as TextBox;
                        TextBox priceTextBox = Controls.Find("txtprice" + i, true).FirstOrDefault() as TextBox;

                        string materialName = comboBox.Text;
                        int quantity = int.Parse(quantityTextBox.Text);
                        int price = int.Parse(priceTextBox.Text);

                        listRun.AppendChild(new Text("Товар: " + materialName));
                        listRun.AppendChild(new Break());
                        listRun.AppendChild(new Text("Количество: " + quantity));
                        listRun.AppendChild(new Break());
                        listRun.AppendChild(new Text("Цена: " + price));
                        listRun.AppendChild(new Break());
                        listRun.AppendChild(new Break());
                    }
                    mainPart.Document.Save();
                }

                MessageBox.Show("Документ успешно сохранен в формате Word.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при сохранении документа: " + ex.Message);
            }
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            word.Text = "Сохранить\n в Word";
        }
    }
}
