using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;//libreria sql
using System.Data.Sql;//libreria sql
namespace Nominas
{
    public partial class Form1 : Form
    {
        DataTable CargarChecadas = new DataTable();
        DataTable CargarTurnos = new DataTable();
        SqlConnection cn; // para la conexion a SQL
        SqlCommand cmd;  // para las instrucciones SQL
        SqlDataReader dr; //
        CadenaConexion var = new CadenaConexion();
        public Form1()
        {
            InitializeComponent();
        }

        private void Pruebas_Click(object sender, EventArgs e)
        {
            this.LimpiarTablasTemporales();//SE LIMPIAN LAS TABLAS ANTES DE COMENZAR EL LLENADO
            /*::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::*/
            this.CargarChecadasTemporal();//Carga el DataSet Checadas a una tabla temporal a SQL
            this.CargarTurnosTemporal();//Carga el DataSet Turnos a una tabla temporal a SQL
           /*::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::*/
            DateTime FechaInicio, FechaFinal;
            String FechaFormateada, FinalFormateada;

            FechaInicio = dateInicio.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro
            FechaFormateada = FechaInicio.ToString("yyyy-MM-dd 00:00:00:000");

            FechaFinal = dateFinal.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro  
            FechaFinal = FechaFinal.AddDays(1);
            FinalFormateada = FechaFinal.ToString("yyyy-MM-dd 00:00:00:000");

            this.InsertarChecadasTemporal(FechaFormateada, FinalFormateada);
            /*::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::*/
            this.RenombrarTurnos();//Cambia los Turnos del Acces a los creados en C#
        }

        private void TblChecadaBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.tblChecadaBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.checadorDataSet);
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            // TODO: esta línea de código carga datos en la tabla 'checadorDataSet.tblTrabTurno' Puede moverla o quitarla según sea necesario.
            this.tblTrabTurnoTableAdapter.Fill(this.checadorDataSet.tblTrabTurno);
            // TODO: esta línea de código carga datos en la tabla 'checadorDataSet.tblChecada' Puede moverla o quitarla según sea necesario.
            this.tblChecadaTableAdapter.Fill(this.checadorDataSet.tblChecada);
            // TODO: esta línea de código carga datos en la tabla 'checadorDataSet.tblAcceso' Puede moverla o quitarla según sea necesario.

            CargarChecadas = checadorDataSet.Tables["tblChecada"];//DataTables Cargados de los DataSets
            CargarTurnos = checadorDataSet.Tables["tblTrabTurno"];//DataTables Cargados de los DataSets
        }


        private void CargarChecadasTemporal()
        {
            String Planta;
            if (rdbBravo.Checked == true)
            {
                Planta = "Bravo";
            }
            else
            {
                Planta = "Morelos";
            }

            for (int i = 0; i < CargarChecadas.Rows.Count; i++)
            {
                try//Elimino la base de datos actual
                {
                    String Trabajador = CargarChecadas.Rows[i]["Trabajador"].ToString();
                    DateTime Checada = Convert.ToDateTime(CargarChecadas.Rows[i]["Checada"].ToString());
                    String TipoChecada = CargarChecadas.Rows[i]["TipoChecada"].ToString();
                    cn = new SqlConnection(var.GetConexionDB());
                    cn.Open();
                    SqlCommand cmd = new SqlCommand("CargarChecadasTemporales", cn);
                    cmd.Parameters.AddWithValue("@Trabajador", Trabajador);
                    cmd.Parameters.AddWithValue("@Checada", Checada);
                    cmd.Parameters.AddWithValue("@TipoChecada", TipoChecada);
                    cmd.Parameters.AddWithValue("@Planta", Planta);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.ExecuteReader();
                    cn.Close();
                }
                catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            }
        }

        private void CargarTurnosTemporal()
        {
            String Planta;
            if (rdbBravo.Checked == true)
            {
                Planta = "Bravo";
            }
            else
            {
                Planta = "Morelos";
            }

            for (int i = 0; i < CargarTurnos.Rows.Count; i++)
            {
                try//Elimino la base de datos actual
                {
                    String Trabajador = CargarTurnos.Rows[i]["Trabajador"].ToString();
                    DateTime Fecha = Convert.ToDateTime(CargarTurnos.Rows[i]["Fecha"].ToString());
                    String Turno = CargarTurnos.Rows[i]["Turno"].ToString();
                    cn = new SqlConnection(var.GetConexionDB());
                    cn.Open();
                    SqlCommand cmd = new SqlCommand("CargarTurnoTemporal", cn);
                    cmd.Parameters.AddWithValue("@Trabajador", Trabajador);
                    cmd.Parameters.AddWithValue("@Fecha", Fecha);
                    cmd.Parameters.AddWithValue("@Turno", Turno);
                    cmd.Parameters.AddWithValue("@Planta", Planta);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.ExecuteReader();
                    cn.Close();
                }
                catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            }
        }

        private void InsertarChecadasTemporal(String Fecha_Inicio, String Fecha_Fin)
        {
            String Planta;
            if (rdbBravo.Checked == true)
            {
                Planta = "Bravo";
            }
            else
            {
                Planta = "Morelos";
            }

            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            SqlCommand cmd = new SqlCommand("UnificarTemporalesDB", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@Planta", Planta);
            cmd.Parameters.AddWithValue("@Fecha_Inicio", Fecha_Inicio);
            cmd.Parameters.AddWithValue("@Fecha_Fin", Fecha_Fin);
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Clear();//Limpia antes de llenar
            dt.Load(cmd.ExecuteReader());
            cn.Close();

            gridSemanaTrabajador.DataSource = dt;
            for (int x = 0; x < dt.Rows.Count; x++)//SE INSERTAN LAS CHECADAS A PARTIR DE LA CONSULTA DE TABLAS TEMPORALES ANTERIORIORES EN LA TABLA CHECADAS
            {
                cn.Open();
                cmd = new SqlCommand("InsertarChecadasFinales", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Trabajador", dt.Rows[x]["Trabajador"].ToString());
                cmd.Parameters.AddWithValue("@Checada", Convert.ToDateTime(dt.Rows[x]["Checada"].ToString()));
                cmd.Parameters.AddWithValue("@TipoChecada", dt.Rows[x]["TipoChecada"].ToString());
                cmd.Parameters.AddWithValue("@Turno", dt.Rows[x]["Turno"].ToString());
                cmd.ExecuteReader();
                cn.Close();
            }

            cn.Open();//CAMBIO LOS HORARIOS ASIGNADOS IBIX POR LOS CREADOS EN C#
            cmd = new SqlCommand("RenombrarTurnos", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.ExecuteReader();
            cn.Close();
        }

        private void LimpiarTablasTemporales()
        {
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            SqlCommand cmd = new SqlCommand("LimpiarTablasTemporales", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            System.Data.DataTable dt = new System.Data.DataTable();
            cmd.ExecuteReader();
            cn.Close();
        }

        private void RenombrarTurnos()
        {
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            SqlCommand cmd = new SqlCommand("RenombrarTurnos", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            System.Data.DataTable dt = new System.Data.DataTable();
            cmd.ExecuteReader();
            cn.Close();
        }







    }
}
