using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;//libreria sql
using System.Data;//libreria sql 
using System.Data.Sql;//libreria sql
using System.Windows.Forms;//libreria para usar los messagebox
using System.Data.OleDb;
using System.Globalization;
using System.Collections;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Http;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Nominas
{
    public partial class ChecadorAdministrativo : Form
    {
        public int xClick = 0, yClick = 0;//Variables para mover formulario sin bordes
        String sql;
        String query;
        SqlConnection cn; // para la conexion a SQL
        SqlCommand cmd;  // para las instrucciones SQL
        SqlDataReader dr; //
        CadenaConexion var = new CadenaConexion();
        OperacionesChecador c = new OperacionesChecador();
        SumarHoras sm = new SumarHoras();
        int NOCHE = 2; //Cantidad de checadas
        int SAB = 2; //Cantidad de checadas
        int TARDE = 2; //Cantidad de checadas
        int TDIA = 4; //Cantidad de checadas
        int TFIJ = 4; //Cantidad de checadas
        int TNO2 = 4; //Cantidad de checadas

        ConexionAcces ConA = new ConexionAcces();
        OleDbConnection cna = new OleDbConnection();
        DataTable CargarChecadas = new DataTable();
        DataTable CargarTurnos = new DataTable();
        OperacionesChecador Op = new OperacionesChecador();
        private const int CS_Sombre = 0x00020000; //Sombra winForms
        Funciones_Checadas FC = new Funciones_Checadas();
        OperacionesBaños OPB = new OperacionesBaños();
        OperacionesRH RH = new OperacionesRH();

        /// <summary>
        /// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// </summary>
        public ChecadorAdministrativo()
        {
            InitializeComponent();
        }

        private void ButVerSemanaPrincipal_Click(object sender, EventArgs e)
        {
            this.MostrarSemanaPrincipal();//Muestra la semana
        }

        private void MostrarSemanaPrincipal()
        {
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            ////////////////////////////////// CONSULTA SEMANA DE LOS TRABAJADORES /////////////////////////////////////////////////////////
            DateTime FechaInicio, FechaFinal;
            String FechaFormateada, FinalFormateada;

            FechaInicio = dateInicio.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro
            FechaFormateada = FechaInicio.ToString("yyyy-MM-dd 00:00:00:000");

            FechaFinal = dateFinal.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro  
            FechaFinal = FechaFinal.AddDays(1);
            FinalFormateada = FechaFinal.ToString("yyyy-MM-dd 00:00:00:000");

            MessageBox.Show("Fecha inicio: " + FechaInicio + " \n A \n fecha final: " + FechaFinal);
            try
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                //importar la libreria de widws form para abrir una ventana 
                SqlCommand cmd = new SqlCommand("FechaNomina_Pro", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FechaInicio", FechaFormateada);
                cmd.Parameters.AddWithValue("@FechaFinal", FinalFormateada);
                cmd.Parameters.AddWithValue("@Planta", listPlantas.Text.ToString());
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(cmd.ExecuteReader());
                GridSemana.DataSource = dt;
                cn.Close();
                MessageBox.Show(dt.Rows.Count.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            this.ColorReporte();
        }

        private void CargarBase_Click(object sender, EventArgs e)
        {
            if (GridSemana.Rows.Count > 1)
            {
                for (int i = GridSemana.Rows.Count - 1; i >= 0; i--)
                {
                    GridSemana.Rows.RemoveAt(i);
                }
            }
            GridSemana.Refresh();
            this.LimpiarTablasTemporales();//SE LIMPIAN LAS TABLAS ANTES DE COMENZAR EL LLENADO
            /*::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::*/
            this.CargarChecadasTemporal();//Carga el DataSet Checadas a una tabla temporal a SQL
            this.CargarTurnosTemporal();//Carga el DataSet Turnos a una tabla temporal a SQL
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
            this.MostrarSemanaPrincipal();//Muestra la semana
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

        private void InsertarChecadasTemporal(String Fecha_Inicio, String Fecha_Fin)
        {
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            SqlCommand cmd = new SqlCommand("UnificarTemporalesDB", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@Planta", listPlantas.Text.ToString());
            cmd.Parameters.AddWithValue("@Fecha_Inicio", Fecha_Inicio);
            cmd.Parameters.AddWithValue("@Fecha_Fin", Fecha_Fin);
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Clear();//Limpia antes de llenar
            dt.Load(cmd.ExecuteReader());
            cn.Close();

            BarraProgreso.Maximum = dt.Rows.Count;
            BarraProgreso.BackColor = Color.White;
            BarraProgreso.ProgressColor1 = Color.CornflowerBlue;
            BarraProgreso.ProgressColor2 = Color.DarkRed;
            BarraProgreso.Visible = true;
            BarraProgreso.ForeColor = Color.Cyan;

            for (int x = 0; x < dt.Rows.Count; x++)//SE INSERTAN LAS CHECADAS A PARTIR DE LA CONSULTA DE TABLAS TEMPORALES ANTERIORIORES EN LA TABLA CHECADAS
            {
                if (0 == x % 10)
                {
                    BarraProgreso.Value = x;
                    BarraProgreso.Update();//Contador de progreso
                    BarraProgreso.Visible = true;
                }

                cn.Open();
                cmd = new SqlCommand("InsertarChecadasFinales", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Trabajador", dt.Rows[x]["Trabajador"].ToString());
                cmd.Parameters.AddWithValue("@Checada", Convert.ToDateTime(dt.Rows[x]["Checada"].ToString()));
                cmd.Parameters.AddWithValue("@TipoChecada", dt.Rows[x]["TipoChecada"].ToString());
                cmd.Parameters.AddWithValue("@Turno", dt.Rows[x]["Turno"].ToString());
                cmd.Parameters.AddWithValue("@Planta", dt.Rows[x]["Planta"].ToString());
                cmd.ExecuteReader();
                cn.Close();
            }

            BarraProgreso.Visible = false;
            cn.Open();//CAMBIO LOS HORARIOS ASIGNADOS IBIX POR LOS CREADOS EN C#
            cmd = new SqlCommand("RenombrarTurnos", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.ExecuteReader();
            cn.Close();
        }

        private void CargarChecadasTemporal()
        {
            BarraProgreso.Maximum = CargarChecadas.Rows.Count;
            BarraProgreso.BackColor = Color.White;
            BarraProgreso.ProgressColor1 = Color.CornflowerBlue;
            BarraProgreso.ProgressColor2 = Color.DarkRed;
            BarraProgreso.Visible = true;
            BarraProgreso.ForeColor = Color.Cyan;
            for (int i = 0; i < CargarChecadas.Rows.Count; i++)
            {
                if (0 == i % 10)
                {
                    BarraProgreso.Value = i;
                    BarraProgreso.Update();//Contador de progreso
                    //BarraProgreso.Visible = true;
                }
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
                    cmd.Parameters.AddWithValue("@Planta", listPlantas.Text.ToString());
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.ExecuteReader();
                    cn.Close();
                }
                catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            }
            BarraProgreso.Visible = false;
        }

        private void CargarTurnosTemporal()
        {
            BarraProgreso.Maximum = CargarTurnos.Rows.Count;
            BarraProgreso.BackColor = Color.White;
            BarraProgreso.ProgressColor1 = Color.CornflowerBlue;
            BarraProgreso.ProgressColor2 = Color.DarkRed;
            BarraProgreso.Visible = true;
            BarraProgreso.ForeColor = Color.Cyan;
            for (int i = 0; i < CargarTurnos.Rows.Count; i++)
            {
                if (0 == i % 10)
                {
                    BarraProgreso.Value = i;
                    BarraProgreso.Update();//Contador de progreso
                    //BarraProgreso.Visible = true;
                }
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
                    cmd.Parameters.AddWithValue("@Planta", listPlantas.Text.ToString());
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.ExecuteReader();
                    cn.Close();
                }
                catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            }
            BarraProgreso.Visible = false;
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

        private void ButFiltrar_Click(object sender, EventArgs e)
        {
            c.CrearTablaRespaldo();
            this.Filtrar();
            RH.ListaDeIncidentes(dateInicio.Value, dateFinal.Value);
            MessageBox.Show("Fin de filtrado");
        }

        private void Filtrar()
        {
            ///////////////////////////////////// SELECCIÓN DE EMPLEADOS/////////////////////////////////////////////////////////////
            ArrayList Empleado = new ArrayList();
            ArrayList Lunes = new ArrayList();
            ArrayList Martes = new ArrayList();
            ArrayList Miercoles = new ArrayList();
            ArrayList Jueves = new ArrayList();
            ArrayList Viernes = new ArrayList();
            ArrayList Sabado = new ArrayList();
            DateTime Falta = new DateTime(2000, 01, 01, 0, 0, 0);
            /////////////////////////////////////GET COLECCIÓN && SELECCIÓN DE EMPLEADOS/////////////////////////////////////////////////////////////
            //Empleado.Clear();
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            query = "SELECT DISTINCT Trabajador FROM Checadas WHERE Turno = 'ADMIN' OR Turno = '' Order by Trabajador ";

            SqlCommand cmd2 = new SqlCommand(query, cn);
            DataTable t1 = new DataTable();
            using (SqlDataAdapter a = new SqlDataAdapter(cmd2))
            {
                a.Fill(t1);
            }
            cn.Close();
            for (int x = 0; x < t1.Rows.Count; x++)//FOR[1] : Agrega los Trabajadores a una lista
            {
                Empleado.Add(t1.Rows[x]["Trabajador"].ToString());
            }
            MessageBox.Show("Lista empleados" + Empleado.Count.ToString() + " Primero" + Empleado[0].ToString() );
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            ////////////////////////////////// CONSULTA SEMANA DE TRABAJADOR /////////////////////////////////////////////////////////

            for (int NumNomina = 0; NumNomina < Empleado.Count; NumNomina++)//FOR[2]: RECORRE TODA LA LISTA DE TRABAJADORES
            {
                Lunes.Clear();
                Martes.Clear();
                Miercoles.Clear();
                Jueves.Clear();
                Viernes.Clear();
                Viernes.Clear();
                Sabado.Clear();

                DateTime FechaInicio, FechaFinal;
                String FechaFormateada, FinalFormateada;

                FechaInicio = dateInicio.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro
                                               // FechaInicio=FechaInicio.AddDays(-1);//Se agrega 1 dia más
                FechaFormateada = FechaInicio.ToString("yyyy-MM-dd 00:00:00:000");

                FechaFinal = dateFinal.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro  
                FechaFinal = FechaFinal.AddDays(1);
                FinalFormateada = FechaFinal.ToString("yyyy-MM-dd 00:00:00:000");

                try//try general
                {
                    DataTable dt = c.SemanaTrabajador(Empleado[NumNomina].ToString(), FechaFormateada, FinalFormateada);
                    MessageBox.Show("Checadas de empleados: " + dt.Rows.Count.ToString() + " Nomina" + Empleado[NumNomina].ToString());
                    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    //////////////////////////////////////////// MOSTRAR SEMANA GRID /////////////////////////////////////////////////////////
                    DateTime dateValue = new DateTime();
                    CultureInfo cul = CultureInfo.CurrentCulture;
                    // Usa la fecha formateada y calcula el número de la semana
                    int NumeroSemana = cul.Calendar.GetWeekOfYear(dateValue, CalendarWeekRule.FirstDay, DayOfWeek.Monday);

                    for (int i = 0; i < dt.Rows.Count; i++)//FOR[3]: Agrega los datos del DT a las listas de dia
                    {
                        dateValue = Convert.ToDateTime(dt.Rows[i]["Checada"].ToString());//Consulto la checada del indice
                        String dia = dateValue.DayOfWeek.ToString();//Convierto la fecha en dia de la semana
                        switch (dia)
                        {
                            case "Lunes":
                            case "Monday":
                                Lunes.Add(dt.Rows[i]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                break;
                            case "Martes":
                            case "Tuesday":
                                Martes.Add(dt.Rows[i]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                break;
                            case "Miércoles":
                            case "Wednesday":
                                Miercoles.Add(dt.Rows[i]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                break;
                            case "Jueves":
                            case "Thursday":
                                Jueves.Add(dt.Rows[i]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                break;
                            case "Viernes":
                            case "Friday":
                                Viernes.Add(dt.Rows[i]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                break;
                            case "Sábado":
                            case "Saturday":
                                Sabado.Add(dt.Rows[i]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                break;
                        }
                    }

                    if ((Lunes.Count <= 4) && (Martes.Count <= 4) && (Miercoles.Count <= 4) && (Jueves.Count <= 4) && (Viernes.Count <= 4) && (Sabado.Count <= 2))
                    {
                        MessageBox.Show("Empleado: " + Empleado[NumNomina].ToString() + " Lunes: " + Lunes.Count.ToString() + "Martes: " + Martes.Count.ToString() + "Miercoles: " + Miercoles.Count.ToString() + "Jueves: " + Jueves.Count.ToString() + "Viernes: " + Viernes.Count.ToString() + "Sabado: " + Sabado.Count.ToString());
                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                    }else
                    {
                        MessageBox.Show("Empleado: " + Empleado[NumNomina].ToString() + " NO CUMPLE");
                       c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                    }
                }
                catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
            }
        }

        private void TblChecadaBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.tblChecadaBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.checadorDataSet);
        }

        private void ChecadorAdministrativo_Load(object sender, EventArgs e)
        {
            // TODO: esta línea de código carga datos en la tabla 'checadorDataSet.tblTrabTurno' Puede moverla o quitarla según sea necesario.
            // this.tblTrabTurnoTableAdapter.Fill(this.checadorDataSet.tblTrabTurno);
            // TODO: esta línea de código carga datos en la tabla 'checadorDataSet.tblChecada' Puede moverla o quitarla según sea necesario.
            //this.tblChecadaTableAdapter.Fill(this.checadorDataSet.tblChecada);

            CargarChecadas = this.GetChecadasAcces();//DataTables Cargados de los DataSets
            CargarTurnos = this.GetTurnosAcces();//DataTables Cargados de los DataSets

            //Diseño del Datagrid
            GridSemana.BorderStyle = BorderStyle.None;
            GridSemana.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238, 239, 249);
            GridSemana.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            GridSemana.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            GridSemana.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            GridSemana.BackgroundColor = Color.White;
            GridSemana.EnableHeadersVisualStyles = false;
            GridSemana.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            GridSemana.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 20, 24);
            GridSemana.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            //Diseño del Datagrid
            gridComparador.BorderStyle = BorderStyle.None;
            gridComparador.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238, 239, 249);
            gridComparador.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            gridComparador.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            gridComparador.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            gridComparador.BackgroundColor = Color.White;
            gridComparador.EnableHeadersVisualStyles = false;
            gridComparador.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            gridComparador.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 20, 24);
            gridComparador.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            //Diseño del Datagrid
            gridSemanaTrabajador.BorderStyle = BorderStyle.None;
            gridSemanaTrabajador.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238, 239, 249);
            gridSemanaTrabajador.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            gridSemanaTrabajador.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            gridSemanaTrabajador.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            gridSemanaTrabajador.BackgroundColor = Color.White;
            gridSemanaTrabajador.EnableHeadersVisualStyles = false;
            gridSemanaTrabajador.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            gridSemanaTrabajador.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 20, 24);
            gridSemanaTrabajador.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            //Diseño del Datagrid
            gridErrores.BorderStyle = BorderStyle.None;
            gridErrores.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238, 239, 249);
            gridErrores.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            gridErrores.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            gridErrores.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            gridErrores.BackgroundColor = Color.White;
            gridErrores.EnableHeadersVisualStyles = false;
            gridErrores.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            gridErrores.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 20, 24);
            gridErrores.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            //Diseño del Datagrid
            GridSemanaHorarios.BorderStyle = BorderStyle.None;
            GridSemanaHorarios.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238, 239, 249);
            GridSemanaHorarios.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            GridSemanaHorarios.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            GridSemanaHorarios.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            GridSemanaHorarios.BackgroundColor = Color.White;
            GridSemanaHorarios.EnableHeadersVisualStyles = false;
            GridSemanaHorarios.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            GridSemanaHorarios.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 20, 24);
            GridSemanaHorarios.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            //Diseño del Datagrid
            gridIncidencias.BorderStyle = BorderStyle.None;
            gridIncidencias.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238, 239, 249);
            gridIncidencias.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            gridIncidencias.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            gridIncidencias.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            gridIncidencias.BackgroundColor = Color.White;
            gridIncidencias.EnableHeadersVisualStyles = false;
            gridIncidencias.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            gridIncidencias.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(235, 63, 63);
            gridIncidencias.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            //Diseño del Datagrid
            gridErroresIncidencias.BorderStyle = BorderStyle.None;
            gridErroresIncidencias.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238, 239, 249);
            gridErroresIncidencias.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            gridErroresIncidencias.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            gridErroresIncidencias.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            gridErroresIncidencias.BackgroundColor = Color.White;
            gridErroresIncidencias.EnableHeadersVisualStyles = false;
            gridErroresIncidencias.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            gridErroresIncidencias.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(235, 63, 63);
            gridErroresIncidencias.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            //Diseño del Datagrid
            gridSemanaDeRetardos.BorderStyle = BorderStyle.None;
            gridSemanaDeRetardos.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238, 239, 249);
            gridSemanaDeRetardos.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            gridSemanaDeRetardos.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            gridSemanaDeRetardos.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            gridSemanaDeRetardos.BackgroundColor = Color.White;
            gridSemanaDeRetardos.EnableHeadersVisualStyles = false;
            gridSemanaDeRetardos.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            gridSemanaDeRetardos.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(235, 63, 63);
            gridSemanaDeRetardos.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            this.SetFechaActual();
        }

        private DataTable GetChecadasAcces()
        {
            cna = new OleDbConnection(ConA.GetConexionAcces());

            cna.Open();
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataSet ds = new DataSet();
            da.SelectCommand = new OleDbCommand("select * FROM tblChecada ", cna);
            da.Fill(ds);
            dt = ds.Tables[0];
            return dt;
        }

        private DataTable GetTurnosAcces()
        {
            cna = new OleDbConnection(ConA.GetConexionAcces());
            cna.Open();
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataSet ds = new DataSet();
            da.SelectCommand = new OleDbCommand("select * FROM tblTrabTurno ", cna);
            da.Fill(ds);
            dt = ds.Tables[0];
            return dt;
        }

        private void Descontar_Click(object sender, EventArgs e)
        {
            RH.LimpiarListaRetardos();
            DateTime EntradaAdmin = Convert.ToDateTime("08:00:00");
            ArrayList Lunes = new ArrayList();
            ArrayList Martes = new ArrayList();
            ArrayList Miercoles = new ArrayList();
            ArrayList Jueves = new ArrayList();
            ArrayList Viernes = new ArrayList();
            ArrayList Sabado = new ArrayList();
            ArrayList Empleado = new ArrayList();

            /////////////////////////////////////GET COLECCIÓN && SELECCIÓN DE EMPLEADOS POR TURNOS /////////////////////////////////////////////////////////////

            Empleado.Clear();
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            string query = "SELECT DISTINCT Trabajador FROM Checadas WHERE Turno = 'ADMIN' OR Turno = '' Order by Trabajador ";

            SqlCommand cmd = new SqlCommand(query, cn);
            DataTable t1 = new DataTable();
            using (SqlDataAdapter a = new SqlDataAdapter(cmd))
            {
                a.Fill(t1);
            }
            cn.Close();

            for (int x = 0; x < t1.Rows.Count; x++)//FOR[1] : Agrega los Trabajadores a una lista
            {
                Empleado.Add(t1.Rows[x]["Trabajador"].ToString());
            }
            ///////////////////////////////////////////// OBTENER DATOS DE CADA EMPLEADO EN LISTA //////////////////////////////////////////////////////
            for (int indiceEmpleados = 0; indiceEmpleados < Empleado.Count; indiceEmpleados++)
            {
                Lunes.Clear();
                Martes.Clear();
                Miercoles.Clear();
                Jueves.Clear();
                Viernes.Clear();
                Viernes.Clear();
                Sabado.Clear();

                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                query = "SELECT * FROM Checadas WHERE Trabajador = '" + Empleado[indiceEmpleados] + "' Order by Trabajador,Checada ";

                SqlCommand cmd2 = new SqlCommand(query, cn);
                DataTable dt = new DataTable();
                using (SqlDataAdapter a = new SqlDataAdapter(cmd2))
                {
                    a.Fill(dt);
                }
                cn.Close();
                /////////////////////////////////////////////// OBTENER CHECADAS DEL DIA /////////////////////////////////////////////////////
                for (int reg = 0; reg < dt.Rows.Count; reg++)
                {
                    DateTime dateValue = new DateTime();
                    CultureInfo cul = CultureInfo.CurrentCulture;
                    // Usa la fecha formateada y calcula el número de la semana
                    int NumeroSemana = cul.Calendar.GetWeekOfYear(dateValue, CalendarWeekRule.FirstDay, DayOfWeek.Monday);

                    dateValue = Convert.ToDateTime(dt.Rows[reg]["Checada"].ToString());//Consulto la checada del indice
                    String dia = dateValue.DayOfWeek.ToString();//Convierto la fecha en dia de la semana
                    switch (dia)
                    {
                        case "Lunes":
                        case "Monday":
                            Lunes.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                            break;
                        case "Martes":
                        case "Tuesday":
                            Martes.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                            break;
                        case "Miércoles":
                        case "Wednesday":
                            Miercoles.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                            break;
                        case "Jueves":
                        case "Thursday":
                            Jueves.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                            break;
                        case "Viernes":
                        case "Friday":
                            Viernes.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                            break;
                        case "Sábado":
                        case "Saturday":
                            Sabado.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                            break;
                    }//FIN SWITCH dias
                }//Fin FOR Checada por dias

                if (Lunes.Count > 0)
                {
                    TimeSpan Admin = new TimeSpan(08, 00, 0);//Variable con el horario 19:00 hrs
                    EntradaAdmin = Convert.ToDateTime(Lunes[0].ToString()).Date + Admin;//Establece fecha de la checada con la hora de ts
                    if (Convert.ToDateTime(Lunes[0].ToString()) < EntradaAdmin.AddMinutes(10))
                    {
                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaAdmin);//Actualiza la checada al turno asignado
                    }
                    else
                    {
                        c.LLegadaTardeAdmin(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaAdmin);
                    }
                }
                if (Martes.Count > 0)
                {
                    TimeSpan Admin = new TimeSpan(08, 00, 0);//Variable con el horario 19:00 hrs
                    EntradaAdmin = Convert.ToDateTime(Martes[0].ToString()).Date + Admin;//Establece fecha de la checada con la hora de ts
                    if (Convert.ToDateTime(Martes[0].ToString()) < EntradaAdmin.AddMinutes(10))
                    {
                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), EntradaAdmin);//Actualiza la checada al turno asignado
                    }
                    else
                    {
                        c.LLegadaTardeAdmin(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), EntradaAdmin);
                    }
                }
                if (Miercoles.Count > 0)
                {
                    TimeSpan Admin = new TimeSpan(08, 00, 0);//Variable con el horario 19:00 hrs
                    EntradaAdmin = Convert.ToDateTime(Miercoles[0].ToString()).Date + Admin;//Establece fecha de la checada con la hora de ts
                    if (Convert.ToDateTime(Miercoles[0].ToString()) < EntradaAdmin.AddMinutes(10))
                    {
                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), EntradaAdmin);//Actualiza la checada al turno asignado
                    }
                    else
                    {
                        c.LLegadaTardeAdmin(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), EntradaAdmin);
                    }
                }
                if (Jueves.Count > 0)
                {
                    TimeSpan Admin = new TimeSpan(08, 00, 0);//Variable con el horario 19:00 hrs
                    EntradaAdmin = Convert.ToDateTime(Jueves[0].ToString()).Date + Admin;//Establece fecha de la checada con la hora de ts
                    if (Convert.ToDateTime(Jueves[0].ToString()) < EntradaAdmin.AddMinutes(10))
                    {
                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), EntradaAdmin);//Actualiza la checada al turno asignado
                    }
                    else
                    {
                        c.LLegadaTardeAdmin(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), EntradaAdmin);
                    }
                }
                if (Viernes.Count > 0)
                {
                    TimeSpan Admin = new TimeSpan(08, 00, 0);//Variable con el horario 19:00 hrs
                    EntradaAdmin = Convert.ToDateTime(Viernes[0].ToString()).Date + Admin;//Establece fecha de la checada con la hora de ts
                    if (Convert.ToDateTime(Viernes[0].ToString()) < EntradaAdmin.AddMinutes(10))
                    {
                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), EntradaAdmin);//Actualiza la checada al turno asignado
                    }
                    else
                    {
                        c.LLegadaTardeAdmin(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), EntradaAdmin);
                    }
                }
                if (Sabado.Count > 0)
                {
                    TimeSpan Admin = new TimeSpan(08, 00, 0);//Variable con el horario 19:00 hrs
                    EntradaAdmin = Convert.ToDateTime(Sabado[0].ToString()).Date + Admin;//Establece fecha de la checada con la hora de ts
                    if (Convert.ToDateTime(Sabado[0].ToString()) < EntradaAdmin.AddMinutes(10))
                    {
                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), EntradaAdmin);//Actualiza la checada al turno asignado
                    }
                    else
                    {
                        c.LLegadaTardeAdmin(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), EntradaAdmin);
                    }
                }
            }//For recorrer lista empleados
        }

        public void ColorReporte()//Colorear Grid
        {
            try
            {
                for (int fila = 0; fila < GridSemana.Rows.Count; fila++)
                {
                    String Estatus = GridSemana.Rows[fila].Cells["Estatus"].Value.ToString();
                    if (Estatus.Equals("Correcto"))
                    {
                        GridSemana.Rows[fila].DefaultCellStyle.BackColor = Color.LightSeaGreen;
                    }
                    else if (Estatus.Equals("Error"))
                    {
                        GridSemana.Rows[fila].DefaultCellStyle.BackColor = Color.Red;
                    }//Fin de condicion
                    else
                    {
                        GridSemana.Rows[fila].DefaultCellStyle.BackColor = Color.GreenYellow;
                    }
                }//Fin de For Recorrido Tabla
            }
            catch (Exception ex)
            {
            }
        }

        private void ButDuplicados_Click(object sender, EventArgs e)
        {
            this.Duplicados();
        }

        private void Pruebas_Click(object sender, EventArgs e)
        {

        }

        private void Duplicados()
        {
            ArrayList Empleado = new ArrayList();
            ArrayList Lunes = new ArrayList();
            ArrayList Martes = new ArrayList();
            ArrayList Miercoles = new ArrayList();
            ArrayList Jueves = new ArrayList();
            ArrayList Viernes = new ArrayList();
            ArrayList Sabado = new ArrayList();
            /////////////////////////////////////GET COLECCIÓN && SELECCIÓN DE EMPLEADOS POR TURNOS /////////////////////////////////////////////////////////////
            Empleado.Clear();
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            string query = "SELECT DISTINCT Trabajador FROM Checadas WHERE Turno = 'ADMIN' OR Turno = '' Order by Trabajador ";

            SqlCommand cmd = new SqlCommand(query, cn);
            DataTable t1 = new DataTable();
            using (SqlDataAdapter a = new SqlDataAdapter(cmd))
            {
                a.Fill(t1);
            }
            cn.Close();

            for (int x = 0; x < t1.Rows.Count; x++)//FOR[1] : Agrega los Trabajadores a una lista
            {
                Empleado.Add(t1.Rows[x]["Trabajador"].ToString());
            }
            ///////////////////////////////////////////// OBTENER DATOS DE CADA EMPLEADO EN LISTA //////////////////////////////////////////////////////
            for (int indiceEmpleados = 0; indiceEmpleados < Empleado.Count; indiceEmpleados++)
            {
                Lunes.Clear();
                Martes.Clear();
                Miercoles.Clear();
                Jueves.Clear();
                Viernes.Clear();
                Viernes.Clear();
                Sabado.Clear();

                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                query = "SELECT * FROM Checadas WHERE Trabajador = '" + Empleado[indiceEmpleados] + "' Order by Trabajador,Checada ";

                SqlCommand cmd2 = new SqlCommand(query, cn);
                DataTable dt = new DataTable();
                using (SqlDataAdapter a = new SqlDataAdapter(cmd2))
                {
                    a.Fill(dt);
                }
                cn.Close();

                for (int reg = 0; reg < dt.Rows.Count; reg++)
                {
                    DateTime dateValue = new DateTime();
                    CultureInfo cul = CultureInfo.CurrentCulture;
                    // Usa la fecha formateada y calcula el número de la semana
                    int NumeroSemana = cul.Calendar.GetWeekOfYear(dateValue, CalendarWeekRule.FirstDay, DayOfWeek.Monday);

                    dateValue = Convert.ToDateTime(dt.Rows[reg]["Checada"].ToString());//Consulto la checada del indice
                    String dia = dateValue.DayOfWeek.ToString();//Convierto la fecha en dia de la semana
                    switch (dia)
                    {
                        case "Lunes":
                        case "Monday":
                            Lunes.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                            break;
                        case "Martes":
                        case "Tuesday":
                            Martes.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                            break;
                        case "Miércoles":
                        case "Wednesday":
                            Miercoles.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                            break;
                        case "Jueves":
                        case "Thursday":
                            Jueves.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                            break;
                        case "Viernes":
                        case "Friday":
                            Viernes.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                            break;
                        case "Sábado":
                        case "Saturday":
                            Sabado.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                            break;
                    }//FIN SWITCH dias
                }//Fin FOR Checada por dias

                int indice = 0;
                while (indice < 7)
                {
                    switch (indice)
                    {
                        case 1:
                            if (Lunes.Count > 0)
                            {
                                for (int x = 0; x < Lunes.Count; x++)
                                {
                                    if (x<Lunes.Count-1)
                                    {
                                        DateTime Pivote = new DateTime();
                                        Pivote = Convert.ToDateTime(Lunes[x].ToString());
                                        var Diferencia = Convert.ToDateTime(Lunes[x + 1].ToString()) - Pivote;
                                        if (Convert.ToDouble(Diferencia.TotalMinutes) <= 2)
                                        {
                                            MessageBox.Show("Empleado: " + Empleado[indiceEmpleados].ToString() + " Checada 2: " + Lunes[x + 1].ToString() + " - " + Lunes[x].ToString() + " = " + Diferencia.TotalMinutes.ToString());

                                        }
                                    }
                                }
                            }
                            break;
                        case 2:
                            if (Martes.Count > 0)
                            {
                                for (int x = 0; x < Martes.Count; x++)
                                {
                                    if (x < Martes.Count - 1)
                                    {
                                        DateTime Pivote = new DateTime();
                                        Pivote = Convert.ToDateTime(Martes[x].ToString());
                                        var Diferencia = Convert.ToDateTime(Martes[x + 1].ToString()) - Pivote;
                                        if (Convert.ToDouble(Diferencia.TotalMinutes) <= 2)
                                        {
                                            MessageBox.Show("Empleado: " + Empleado[indiceEmpleados].ToString() + " Checada 2: " + Martes[x + 1].ToString() + " - " + Martes[x].ToString() + " = " + Diferencia.TotalMinutes.ToString());

                                        }
                                    }
                                }
                            }
                            break;
                        case 3:
                            if (Miercoles.Count > 0)
                            {
                                for (int x = 0; x < Miercoles.Count; x++)
                                {
                                    if (x < Miercoles.Count - 1)
                                    {
                                        DateTime Pivote = new DateTime();
                                        Pivote = Convert.ToDateTime(Miercoles[x].ToString());
                                        var Diferencia = Convert.ToDateTime(Miercoles[x + 1].ToString()) - Pivote;
                                        if (Convert.ToDouble(Diferencia.TotalMinutes) <= 2)
                                        {
                                            MessageBox.Show("Empleado: " + Empleado[indiceEmpleados].ToString() + " Checada 2: " + Miercoles[x + 1].ToString() + " - " + Miercoles[x].ToString() + " = " + Diferencia.TotalMinutes.ToString());

                                        }
                                    }
                                }
                            }
                            break;
                        case 4:
                            if (Jueves.Count > 0)
                            {
                                for (int x = 0; x < Jueves.Count; x++)
                                {
                                    if (x < Jueves.Count - 1)
                                    {
                                        DateTime Pivote = new DateTime();
                                        Pivote = Convert.ToDateTime(Jueves[x].ToString());
                                        var Diferencia = Convert.ToDateTime(Jueves[x + 1].ToString()) - Pivote;
                                        if (Convert.ToDouble(Diferencia.TotalMinutes) <= 2)
                                        {
                                            MessageBox.Show("Empleado: " + Empleado[indiceEmpleados].ToString() + " Checada 2: " + Jueves[x + 1].ToString() + " - " + Jueves[x].ToString() + " = " + Diferencia.TotalMinutes.ToString());

                                        }
                                    }
                                }
                            }
                            break;
                        case 5:
                            if (Viernes.Count > 0)
                            {
                                for (int x = 0; x < Viernes.Count; x++)
                                {
                                    if (x < Viernes.Count - 1)
                                    {
                                        DateTime Pivote = new DateTime();
                                        Pivote = Convert.ToDateTime(Viernes[x].ToString());
                                        var Diferencia = Convert.ToDateTime(Viernes[x + 1].ToString()) - Pivote;
                                        if (Convert.ToDouble(Diferencia.TotalMinutes) <= 2)
                                        {
                                            MessageBox.Show("Empleado: " + Empleado[indiceEmpleados].ToString() + " Checada 2: " + Viernes[x + 1].ToString() + " - " + Viernes[x].ToString() + " = " + Diferencia.TotalMinutes.ToString());

                                        }
                                    }
                                }
                            }
                            break;
                        case 6:
                            if (Sabado.Count > 0)
                            {
                                for (int x = 0; x < Sabado.Count; x++)
                                {
                                    if (x < Sabado.Count - 1)
                                    {
                                        DateTime Pivote = new DateTime();
                                        Pivote = Convert.ToDateTime(Sabado[x].ToString());
                                        var Diferencia = Convert.ToDateTime(Sabado[x + 1].ToString()) - Pivote;
                                        if (Convert.ToDouble(Diferencia.TotalMinutes) <= 2)
                                        {
                                            MessageBox.Show("Empleado: " + Empleado[indiceEmpleados].ToString() + " Checada 2: " + Sabado[x + 1].ToString() + " - " + Sabado[x].ToString() + " = " + Diferencia.TotalMinutes.ToString());

                                        }
                                    }
                                }
                            }
                            break;
                    }//SWITCH
                    indice++;
                }//WHILE  
            }//FOR EMPLEADOS
        }//Fin de duplicados

        private void SetFechaActual()
        {
            dateInicio.Value = DateTime.Now;
            dateFinal.Value = DateTime.Now;
            FechaInicio_Date.Value = DateTime.Now;
            FechaFinal_Date.Value = DateTime.Now;
            dateChecadaNueva.Value = DateTime.Now;
            FechaInicioIncidencia.Value = DateTime.Now;
            FinFechaIncidencia.Value = DateTime.Now;
            TimesHorarioInicio.Value = DateTime.Now;
            TimesHorarioFinal.Value = DateTime.Now;
            InicioComparador.Value = DateTime.Now;
            FinComparador.Value = DateTime.Now;
            CalcularInicio.Value = DateTime.Now;
            CalcularFin.Value = DateTime.Now;
        }


    }
}













