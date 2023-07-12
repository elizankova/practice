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
using System.Collections;
using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace It
{
    public partial class Form3 : Form
    {
        DataBase dataBase = new DataBase();
        string connectionString = "Data Source = ELIZZANKOVA\\SQLEXPRESS; Initial Catalog=Practice; Integrated Security=True";
        public Form3()
        {
            InitializeComponent();
            FillComboBox();
        }


        private void Clear()
        {
            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox3.Text = "";
            comboBox4.Text = "";
            txtDocumentNumber.Text = "";
            qua1.Text = "";
            qua2.Text = "";
            qua3.Text = "";
            qua4.Text = "";
            price1.Text = "";
            price2.Text = "";
            price3.Text = "";
            price4.Text = "";
        }

        private void FillComboBox()
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    string query = "SELECT Наименование FROM Materials";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string materialName = reader.GetString(0);
                            comboBox1.Items.Add(materialName);
                            comboBox2.Items.Add(materialName);
                            comboBox3.Items.Add(materialName);
                            comboBox4.Items.Add(materialName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при получении данных: " + ex.Message);
            }
        }




        private void Form3_Load(object sender, EventArgs e)
        { 
            this.materialsTableAdapter.Fill(this.practiceDataSet.Materials);
            word.Text = "Сохранить\n в Word";

        }

        private void btn_save_Click(object sender, EventArgs e)
        {
            string documentNumber = txtDocumentNumber.Text;
            DateTime documentDate = dateTimePickerDate.Value.Date;

            try
            {
                dataBase.openConnection();
                for (int i = 0; i < 4; i++)
                {
                    ComboBox comboBox = Controls.Find("comboBox" + (i + 1), true)[0] as ComboBox;
                    TextBox quantityTextBox = Controls.Find("qua" + (i + 1), true)[0] as TextBox;
                    TextBox priceTextBox = Controls.Find("price" + (i + 1), true)[0] as TextBox;

                    string materialName = comboBox.SelectedItem.ToString();
                    int quantity = Convert.ToInt32(quantityTextBox.Text);
                    decimal price = Convert.ToDecimal(priceTextBox.Text);

                    string checkQuery = $"SELECT COUNT(*) FROM Sklad WHERE Наименование = '{materialName}' AND Цена = '{price}'";
                    SqlCommand checkCommand = new SqlCommand(checkQuery, dataBase.getConnection());
                    int existingCount = (int)checkCommand.ExecuteScalar();

                    if (existingCount > 0)
                    {
                        // Если товар уже существует, обновляется количество
                        string updateQuery = $"UPDATE Sklad SET Количество = Количество + '{quantity}' WHERE Наименование = '{materialName}' AND Цена = '{price}'";
                        SqlCommand updateCommand = new SqlCommand(updateQuery, dataBase.getConnection());
                        updateCommand.ExecuteNonQuery();
                    }
                    else
                    {
                        // Если товар не существует, добавлятся в таблицу Sklad
                        string insertQuery = $"INSERT INTO Sklad (Id_Материала, [Наименование], Единица_измерения, Количество, Цена) " +
                            $"SELECT Id_Материала, '{materialName}', Единица_измерения, '{quantity}', '{price}'" +
                            $"FROM Materials WHERE Наименование = '{materialName}'";

                        SqlCommand insertCommand = new SqlCommand(insertQuery, dataBase.getConnection());
                        insertCommand.ExecuteNonQuery();
                    }

                    string incomeQuery = $"INSERT INTO Income (Id_Товара, Id_Материала, Наименование, Количество, Цена, [Дата прихода], [Номер документа прихода]) " +
                        $"SELECT ID_Товара, ID_Материала, '{materialName}', '{quantity}', '{price}', '{documentDate}', '{documentNumber}' " +
                        $"FROM Sklad WHERE Наименование = '{materialName}' AND Цена = '{price}'";

                    SqlCommand incomeCommand = new SqlCommand(incomeQuery, dataBase.getConnection());
                    incomeCommand.ExecuteNonQuery();
                }

                MessageBox.Show("Данные успешно сохранены.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при сохранении данных: " + ex.Message);
            }
            dataBase.closeConnection();
        }






        private void btn_cl_Click(object sender, EventArgs e)
        {
            Clear();
        }

        private void word_Click(object sender, EventArgs e)
        {

            try
            {
                string fileName = "Приход.docx";
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
                    titleRun.AppendChild(new Text("Дата документа: " + dateTimePickerDate.Value.ToShortDateString()));
                    titleRun.AppendChild(new Break());
                    titleRun.AppendChild(new Break());

                    // Добавление списка товаров
                    Paragraph listParagraph = body.AppendChild(new Paragraph());
                    Run listRun = listParagraph.AppendChild(new Run());
                    listRun.AppendChild(new Text("Список товаров:"));
                    listRun.AppendChild(new Break());
                    listRun.AppendChild(new Break());

                    for (int i = 0; i < 4; i++)
                    {
                        ComboBox comboBox = Controls.Find("comboBox" + (i + 1), true)[0] as ComboBox;
                        TextBox quantityTextBox = Controls.Find("qua" + (i + 1), true)[0] as TextBox;
                        TextBox priceTextBox = Controls.Find("price" + (i + 1), true)[0] as TextBox;

                        string materialName = comboBox.SelectedItem.ToString();
                        int quantity = Convert.ToInt32(quantityTextBox.Text);
                        decimal price = Convert.ToDecimal(priceTextBox.Text);

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
    }
}
    
 