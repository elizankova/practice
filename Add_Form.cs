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
    public partial class Form2 : Form
    {


        DataBase dataBase = new DataBase();
        public Form2()
        {
            InitializeComponent();
        }

        private void btn_addnew_Click(object sender, EventArgs e)
        {
            dataBase.openConnection();

            var addname = txtbx_addname.Text;
            var addedizm = txtbx_addedizm.Text;

            var addQuery = $"insert into Materials (Наименование, Единица_измерения) values ('{addname}', '{addedizm}')";
            var command = new SqlCommand(addQuery, dataBase.getConnection());
            command.ExecuteNonQuery();
            MessageBox.Show("Запись успешно создана!", "Успех", MessageBoxButtons.OK);
            dataBase.closeConnection();
        }

        private void btn_addclear_Click(object sender, EventArgs e)
        {

        }
    }
}
