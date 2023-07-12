using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace It
{
    public partial class Form5 : Form
    {

        string connectionString = "Data Source = ELIZZANKOVA\\SQLEXPRESS; Initial Catalog=Practice; Integrated Security=True";
        public Form5()
        {
            InitializeComponent();

        }

        private void Form5_Load(object sender, EventArgs e)
        {
            this.skladTableAdapter.Fill(this.practiceDataSet.Sklad);
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private void GenerateReport_Click(object sender, EventArgs e)
        {
            try
            {
                string query = @"SELECT S.Наименование, S.Количество, S.Цена, (S.Количество * S.Цена) AS Сумма
                         FROM Sklad S
                         LEFT JOIN Outcome O ON S.Id_Товара = O.Id_Товара
                         LEFT JOIN Income I ON S.Id_Товара = I.Id_Товара
                         WHERE (O.[Дата расхода] <= @selectedDate OR O.[Дата расхода] IS NULL)
                         AND (I.[Дата прихода] <= @selectedDate OR I.[Дата прихода] IS NULL)
                         ORDER BY S.Наименование";

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    SqlCommand command = new SqlCommand(query, connection);
                    command.Parameters.AddWithValue("@selectedDate", datePicker.Value.Date);

                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    DataTable reportTable = new DataTable();
                    adapter.Fill(reportTable);

                    dataGridView.DataSource = reportTable;

                    decimal totalSum = GetTotalSum(reportTable);
                    txtTotalSum.Text = totalSum.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при генерации отчета: " + ex.Message);
            }
        }

        private decimal GetTotalSum(DataTable reportTable)
        {
            decimal totalSum = 0;

            foreach (DataRow row in reportTable.Rows)
            {
                if (row["Количество"] != DBNull.Value && row["Цена"] != DBNull.Value)
                {
                    int quantity = Convert.ToInt32(row["Количество"]);
                    int price = Convert.ToInt32(row["Цена"]);
                    decimal totalPrice = quantity * price;
                    totalSum += totalPrice;
                }
            }

            return totalSum;
        }
        private void ExportToExcel()
        {
            try
            {
                // Создание нового пакета Excel
                using (ExcelPackage package = new ExcelPackage())
                {
                    // Создание листа
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Отчет");

                    // Заполнение заголовков столбцов
                    for (int col = 1; col <= dataGridView.Columns.Count; col++)
                    {
                        worksheet.Cells[1, col].Value = dataGridView.Columns[col - 1].HeaderText;
                        worksheet.Cells[1, col].Style.Font.Bold = true;
                        worksheet.Cells[1, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    // Заполнение данных из DataGridView
                    for (int row = 0; row < dataGridView.Rows.Count; row++)
                    {
                        for (int col = 0; col < dataGridView.Columns.Count; col++)
                        {
                            worksheet.Cells[row + 2, col + 1].Value = dataGridView.Rows[row].Cells[col].Value;
                        }
                    }

                    // Автоматическое изменение размеров столбцов
                    worksheet.Cells.AutoFitColumns();

                    // Сохранение файла Excel
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Excel Files|*.xlsx";
                    saveFileDialog.Title = "Сохранить отчет";
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                        package.SaveAs(excelFile);
                        MessageBox.Show("Отчет успешно сохранен в файл Excel.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при экспорте отчета в Excel: " + ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }
    }
}
