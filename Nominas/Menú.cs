using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Nominas
{
    public partial class Menú : Form
    {
        ConexionAcces ConA = new ConexionAcces();
        OleDbConnection cna = new OleDbConnection();
        public Menú()
        {
            InitializeComponent();
        }

        private void ButProduccion_Click(object sender, EventArgs e)
        {
            Checador frm = new Checador();
            frm.Show();
        }

        private void ButAdministracion_Click(object sender, EventArgs e)
        {
            ChecadorAdministrativo frm = new ChecadorAdministrativo();
            frm.Show();
        }

        /*
        private void PictureBox1_Click(object sender, EventArgs e)
        {
            cna = new OleDbConnection(ConA.GetConexionAcces());
           
            cna.Open();
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataSet ds = new DataSet();
            da.SelectCommand = new OleDbCommand("select * FROM tblChecada ", cna);
            da.Fill(ds);
            dt = ds.Tables[0];
            MessageBox.Show(dt.Rows.Count.ToString());
        }*/


    }
}
