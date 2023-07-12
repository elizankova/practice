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


namespace It
{

    enum RowState
    {
        Existed,
        New,
        Modified,
        Deleted,
        ModifiedNew
    }

    public partial class Form1 : Form
    {

        DataBase dataBase = new DataBase();
        int selectedRow;

        public Form1()
        {
            InitializeComponent();
        }


        private void CreateColumns()
        {
            dataGridView1.Columns.Add("ID_Материала", "ID");
            dataGridView1.Columns.Add("Наименование", "Наименование");
            dataGridView1.Columns.Add("Единица_измерения", "Ед. изм.");
            dataGridView1.Columns.Add("IsNew", String.Empty);
            dataGridView1.Columns[3].Visible = false;

        }

        private void ReadSingleRow(DataGridView dgw, IDataRecord record)
        {
            dgw.Rows.Add(record.GetInt32(0), record.GetString(1), record.GetString(2), RowState.ModifiedNew);
        }

        private void deleteRow()
        {
            int index = dataGridView1.CurrentCell.RowIndex;


            dataGridView1.Rows[index].Visible = false;
            dataGridView1.Rows[index].Cells[3].Value = RowState.Deleted;
                return;
            
        }

        private void Update()
        {
            dataBase.openConnection();

            for(int index = 0; index < dataGridView1.Rows.Count; index++)
            {
                var rowState = (RowState)dataGridView1.Rows[index].Cells[3].Value;

                if (rowState == RowState.Existed)
                    continue;

                if(rowState == RowState.Deleted)
                {
                    var id = Convert.ToInt32(dataGridView1.Rows[index].Cells[0].Value);
                    var deleteQuery = $"delete from Materials where ID_Материала = '{id}'";
                    var command = new SqlCommand(deleteQuery, dataBase.getConnection());
                    command.ExecuteNonQuery();
                    
                }

                if(rowState == RowState.Modified)
                {
                    var id = dataGridView1.Rows[index].Cells[0].Value.ToString();
                    var name = dataGridView1.Rows[index].Cells[1].Value.ToString();
                    var edizm = dataGridView1.Rows[index].Cells[2].Value.ToString();

                    var changeQuery = $"update Materials set Наименование = '{name}', Единица_измерения = '{edizm}' where ID_Материала = '{id}'";
                    var command = new SqlCommand(changeQuery, dataBase.getConnection());
                    command.ExecuteNonQuery();
                }
            }
            dataBase.closeConnection();
        }
         
        private void ClearFields()
        {
            tb_ID.Text = "";
            tb_name.Text = "";
            tb_edizm.Text = "";

        }
         

        private void RefreshDataGrid(DataGridView dgw)
        {
            dgw.Rows.Clear();

            string QueryString = $"Select * from Materials";
            SqlCommand command = new SqlCommand(QueryString, dataBase.getConnection());


            dataBase.openConnection();

            SqlDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                ReadSingleRow(dgw, reader);     
            }
            reader.Close();
        }





        private void Form1_Load(object sender, EventArgs e)
        {
            CreateColumns();
            RefreshDataGrid(dataGridView1);


        }


        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            selectedRow = e.RowIndex;
            if(e.RowIndex>=0)
            {
                DataGridViewRow row = dataGridView1.Rows[selectedRow];
                tb_ID.Text = row.Cells[0].Value.ToString();
                tb_name.Text = row.Cells[1].Value.ToString();
                tb_edizm.Text = row.Cells[2].Value.ToString();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            RefreshDataGrid(dataGridView1);
            ClearFields();  
        }

        private void new_btn_Click(object sender, EventArgs e)
        {
            Form2 addfrm = new Form2();
            addfrm.ShowDialog();
        }

        private void delete_btn_Click(object sender, EventArgs e)
        {

            deleteRow();
            ClearFields();

        }

        private void save_btn_Click(object sender, EventArgs e)
        {
            Update();
            ClearFields();
        }


        private void Change()
        {
            var selectedRowIndex = dataGridView1.CurrentCell.RowIndex;

            var id = tb_ID.Text;
            var name = tb_name.Text;
            var edizm = tb_edizm.Text;

            dataGridView1.Rows[selectedRowIndex].SetValues(id, name, edizm);
            dataGridView1.Rows[selectedRowIndex].Cells[3].Value = RowState.Modified;

        }


        private void change_btn_Click(object sender, EventArgs e)
        {
            Change();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ClearFields();
        }
    }
}
