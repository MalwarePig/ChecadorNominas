using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;//libreria sql
using System.Data.Sql;//libreria sql
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
    public partial class Checador : Form
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

        public Checador()
        {
            InitializeComponent();
        }

        protected override CreateParams CreateParams//Sombre debajo Forms
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ClassStyle |= CS_Sombre;
                return cp;
            }
        }

        private void materialDivider1_Click(object sender, EventArgs e)
        {
        }

        public void LLenarGrid(String archivo, String hoja)
        {
            GridSemana.DataSource = c.gridExcel(archivo, hoja);
        }

        private void Checador_Load(object sender, EventArgs e)
        {
            /*
            // TODO: esta línea de código carga datos en la tabla 'checadorDataSet.tblTrabTurno' Puede moverla o quitarla según sea necesario.
            this.tblTrabTurnoTableAdapter.Fill(this.checadorDataSet.tblTrabTurno);
            // TODO: esta línea de código carga datos en la tabla 'checadorDataSet.tblChecada' Puede moverla o quitarla según sea necesario.
            this.tblChecadaTableAdapter.Fill(this.checadorDataSet.tblChecada);
            */
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
            gridTotal.BorderStyle = BorderStyle.None;
            gridTotal.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238, 239, 249);
            gridTotal.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            gridTotal.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            gridTotal.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            gridTotal.BackgroundColor = Color.White;
            gridTotal.EnableHeadersVisualStyles = false;
            gridTotal.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            gridTotal.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(235, 63, 63);
            gridTotal.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

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
            gridBaños.BorderStyle = BorderStyle.None;
            gridBaños.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238, 239, 249);
            gridBaños.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            gridBaños.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            gridBaños.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            gridBaños.BackgroundColor = Color.White;
            gridBaños.EnableHeadersVisualStyles = false;
            gridBaños.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            gridBaños.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 20, 24);
            gridBaños.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

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

            //Diseño del Datagrid
            gridBlackList.BorderStyle = BorderStyle.None;
            gridBlackList.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238, 239, 249);
            gridBlackList.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            gridBlackList.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            gridBlackList.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            gridBlackList.BackgroundColor = Color.White;
            gridBlackList.EnableHeadersVisualStyles = false;
            gridBlackList.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            gridBlackList.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(235, 63, 63);
            gridBlackList.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            //Diseño del Datagrid
            gridSemanaTrabajador.BorderStyle = BorderStyle.None;
            gridSemanaTrabajador.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238, 239, 249);
            gridSemanaTrabajador.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            gridSemanaTrabajador.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            gridSemanaTrabajador.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            gridSemanaTrabajador.BackgroundColor = Color.White;
            gridSemanaTrabajador.EnableHeadersVisualStyles = false;
            gridSemanaTrabajador.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            gridSemanaTrabajador.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(235, 63, 63);
            gridSemanaTrabajador.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            this.FechaActual();//Establece los DateTimePiker a la fecha actual
        }

        private DataTable GetChecadasAcces()
        {
            DateTime FechaInicio;
            String FechaFormateada;
            FechaInicio = dateInicio.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro
            FechaFormateada = FechaInicio.ToString("yyyy-MM-dd");
            cna = new OleDbConnection(ConA.GetConexionAcces());

            cna.Open();
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataSet ds = new DataSet();
            //da.SelectCommand = new OleDbCommand("SELECT * FROM tblChecada", cna);
            da.SelectCommand = new OleDbCommand("SELECT * FROM tblChecada WHERE TipoChecada <> NO AND Checada > " + FechaFormateada + "", cna);
            //da.SelectCommand = new OleDbCommand("SELECT * FROM tblChecada WHERE TipoChecada = NO", cna);
            da.Fill(ds);
            dt = ds.Tables[0];
            return dt;
        }

        private DataTable GetTurnosAcces()
        {
            DateTime FechaInicio;
            String FechaFormateada;
            FechaInicio = dateInicio.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro
            FechaFormateada = FechaInicio.ToString("yyyy-MM-dd");

            cna = new OleDbConnection(ConA.GetConexionAcces());
            cna.Open();
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter();
            DataSet ds = new DataSet();
            da.SelectCommand = new OleDbCommand("SELECT * FROM tblTrabTurno WHERE Fecha > " + FechaFormateada + "", cna);
            //da.SelectCommand = new OleDbCommand("SELECT * FROM tblTrabTurno WHERE Turno = 'NOCHE' OR Turno = 'SAB'  OR Turno = 'TARDE' OR Turno = 'TDIA' OR Turno = 'TFIJ' OR Turno = 'TNO2'", cna);

            da.Fill(ds);
            dt = ds.Tables[0];
            return dt;
        }

        private void butFiltrar_Click(object sender, EventArgs e)
        {
            if (GridSemana.Rows.Count > 1)
            {
                for (int i = GridSemana.Rows.Count - 1; i >= 0; i--)
                {
                    GridSemana.Rows.RemoveAt(i);
                }
            }
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            cmd = new SqlCommand("TRUNCATE TABLE ListaIncidencias", cn);
            cmd.ExecuteNonQuery();//para que relice la sentencia sql
            cn.Close();

            GridSemana.Refresh();
            c.CrearTablaRespaldo();
            ArrayList Turno = new ArrayList() { "NOCHE", "SAB", "TARDE", "TDIA", "TFIJ", "TNO2" };
            //////////////////////////////////// FILTRAR POR HORARIO ////////////////////////////////////////////////////////////////////////////////
            int iTurno = 0;

            BarraProgreso.Maximum = 6;
            BarraProgreso.BackColor = Color.White;
            BarraProgreso.ProgressColor1 = Color.CornflowerBlue;
            BarraProgreso.ProgressColor2 = Color.DarkRed;
            BarraProgreso.Visible = true;
            BarraProgreso.ForeColor = Color.Cyan;
            while (iTurno < 6)
            {
                BarraProgreso.Value = iTurno;
                BarraProgreso.Update();//Contador de progreso
                this.Filtrar(Turno[iTurno].ToString());
                iTurno = iTurno + 1;
            }//WHILE TURNOS

            RH.ListaDeIncidentes(dateInicio.Value, dateFinal.Value);
            BarraProgreso.Visible = false;
            MessageBox.Show("Fin de filtrado");
            chekFiltrar.Checked = true;
        }//FOR FIN DE BOTON

        public void HorasTrabajadas()//Sumador
        {
            DateTime CeldaHora, ET, SC, EC, ST, EI, SI, TOTAL, DIA;
            ArrayList Horas = new ArrayList();
            /*::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
             *:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: LUNES :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::*/
            for (int fila = 0; fila < 6; fila++)//Recorre el lunes
            {
                if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("ET") && GridSemana.Rows[fila].Cells["Lunes"].Value != null)
                {
                    ET = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Lunes"].Value.ToString());
                    MessageBox.Show("Estoy en fila ET: " + ET.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Lunes"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("SC") && GridSemana.Rows[fila].Cells["Lunes"].Value != null)
                {
                    SC = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Lunes"].Value.ToString());
                    MessageBox.Show("Estoy en fila SC: " + SC.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Lunes"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("EC") && GridSemana.Rows[fila].Cells["Lunes"].Value != null)
                {
                    EC = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Lunes"].Value.ToString());
                    MessageBox.Show("Estoy en fila EC: " + EC.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Lunes"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("ST") && GridSemana.Rows[fila].Cells["Lunes"].Value != null)
                {
                    ST = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Lunes"].Value.ToString());
                    MessageBox.Show("Estoy en fila ST: " + ST.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Lunes"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("EI") && GridSemana.Rows[fila].Cells["Lunes"].Value != null)
                {
                    EI = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Lunes"].Value.ToString());
                    MessageBox.Show("Estoy en fila EI: " + EI.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Lunes"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("SI") && GridSemana.Rows[fila].Cells["Lunes"].Value != null)
                {
                    SI = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Lunes"].Value.ToString());
                    MessageBox.Show("Estoy en fila SI: " + SI.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Lunes"].Value.ToString()));
                }
            }//For Lunes
            //ne que poca madre entonces no puedo ento
            TimeSpan Total = Convert.ToDateTime(Horas[1].ToString()) - Convert.ToDateTime(Horas[0].ToString());
            Total = Total + (Convert.ToDateTime(Horas[3].ToString()) - Convert.ToDateTime(Horas[2].ToString()));
            // Difference in days.
            int Lunes = Total.Hours;
            MessageBox.Show(" Son: " + Lunes.ToString() + " en Lunes");
            /*::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::*/
            /*::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            :::::::::::::::::::::::::::::::::::::::::::: MARTES :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::*/
            for (int fila = 0; fila < 6; fila++)//Recorre el Martes
            {
                if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("ET") && GridSemana.Rows[fila].Cells["Martes"].Value != null)
                {
                    ET = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Martes"].Value.ToString());
                    MessageBox.Show("Estoy en fila ET: " + ET.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Martes"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("SC") && GridSemana.Rows[fila].Cells["Martes"].Value != null)
                {
                    SC = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Martes"].Value.ToString());
                    MessageBox.Show("Estoy en fila SC: " + SC.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Martes"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("EC") && GridSemana.Rows[fila].Cells["Martes"].Value != null)
                {
                    EC = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Martes"].Value.ToString());
                    MessageBox.Show("Estoy en fila EC: " + EC.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Martes"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("ST") && GridSemana.Rows[fila].Cells["Martes"].Value != null)
                {
                    ST = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Martes"].Value.ToString());
                    MessageBox.Show("Estoy en fila ST: " + ST.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Martes"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("EI") && GridSemana.Rows[fila].Cells["Martes"].Value != null)
                {
                    EI = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Martes"].Value.ToString());
                    MessageBox.Show("Estoy en fila EI: " + EI.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Martes"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("SI") && GridSemana.Rows[fila].Cells["Martes"].Value != null)
                {
                    SI = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Martes"].Value.ToString());
                    MessageBox.Show("Estoy en fila SI: " + SI.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Martes"].Value.ToString()));
                }
            }//For Martes

            Total = Convert.ToDateTime(Horas[1].ToString()) - Convert.ToDateTime(Horas[0].ToString());
            Total = Total + (Convert.ToDateTime(Horas[3].ToString()) - Convert.ToDateTime(Horas[2].ToString()));
            // Difference in days.
            int Martes = Total.Hours;
            MessageBox.Show(" Son: " + Martes.ToString() + " en Martes");
            /*::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            :::::::::::::::::::::::::::::::::::::::::::: MIERCOLES :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::*/
            for (int fila = 0; fila < 6; fila++)//Recorre el Martes
            {
                if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("ET") && GridSemana.Rows[fila].Cells["Miércoles"].Value != null)
                {
                    ET = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Miércoles"].Value.ToString());
                    MessageBox.Show("Estoy en fila ET: " + ET.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Miércoles"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("SC") && GridSemana.Rows[fila].Cells["Miércoles"].Value != null)
                {
                    SC = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Miércoles"].Value.ToString());
                    MessageBox.Show("Estoy en fila SC: " + SC.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Miércoles"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("EC") && GridSemana.Rows[fila].Cells["Miércoles"].Value != null)
                {
                    EC = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Miércoles"].Value.ToString());
                    MessageBox.Show("Estoy en fila EC: " + EC.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Miércoles"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("ST") && GridSemana.Rows[fila].Cells["Miércoles"].Value != null)
                {
                    ST = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Miércoles"].Value.ToString());
                    MessageBox.Show("Estoy en fila ST: " + ST.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Miércoles"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("EI") && GridSemana.Rows[fila].Cells["Miércoles"].Value != null)
                {
                    EI = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Miércoles"].Value.ToString());
                    MessageBox.Show("Estoy en fila EI: " + EI.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Miércoles"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("SI") && GridSemana.Rows[fila].Cells["Miércoles"].Value != null)
                {
                    SI = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Miércoles"].Value.ToString());
                    MessageBox.Show("Estoy en fila SI: " + SI.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Miércoles"].Value.ToString()));
                }
            }//For Miercoles

            Total = Convert.ToDateTime(Horas[1].ToString()) - Convert.ToDateTime(Horas[0].ToString());
            Total = Total + (Convert.ToDateTime(Horas[3].ToString()) - Convert.ToDateTime(Horas[2].ToString()));
            // Difference in days.
            int Miércoles = Total.Hours;
            MessageBox.Show(" Son: " + Miércoles.ToString() + " en Miércoles");
            /*::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            :::::::::::::::::::::::::::::::::::::::::::: JUEVES :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::*/
            for (int fila = 0; fila < 6; fila++)//Recorre el Martes
            {
                if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("ET") && GridSemana.Rows[fila].Cells["Jueves"].Value != null)
                {
                    ET = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Jueves"].Value.ToString());
                    MessageBox.Show("Estoy en fila ET: " + ET.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Jueves"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("SC") && GridSemana.Rows[fila].Cells["Jueves"].Value != null)
                {
                    SC = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Jueves"].Value.ToString());
                    MessageBox.Show("Estoy en fila SC: " + SC.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Jueves"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("EC") && GridSemana.Rows[fila].Cells["Jueves"].Value != null)
                {
                    EC = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Jueves"].Value.ToString());
                    MessageBox.Show("Estoy en fila EC: " + EC.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Jueves"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("ST") && GridSemana.Rows[fila].Cells["Jueves"].Value != null)
                {
                    ST = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Jueves"].Value.ToString());
                    MessageBox.Show("Estoy en fila ST: " + ST.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Jueves"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("EI") && GridSemana.Rows[fila].Cells["Jueves"].Value != null)
                {
                    EI = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Jueves"].Value.ToString());
                    MessageBox.Show("Estoy en fila EI: " + EI.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Jueves"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("SI") && GridSemana.Rows[fila].Cells["Jueves"].Value != null)
                {
                    SI = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Jueves"].Value.ToString());
                    MessageBox.Show("Estoy en fila SI: " + SI.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Jueves"].Value.ToString()));
                }
            }//For Jueves

            Total = Convert.ToDateTime(Horas[1].ToString()) - Convert.ToDateTime(Horas[0].ToString());
            Total = Total + (Convert.ToDateTime(Horas[3].ToString()) - Convert.ToDateTime(Horas[2].ToString()));
            // Difference in days.
            int Jueves = Total.Hours;
            MessageBox.Show(" Son: " + Jueves.ToString() + " en Jueves");
            /*::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
            :::::::::::::::::::::::::::::::::::::::::::: Viernes :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::*/
            for (int fila = 0; fila < 6; fila++)//Recorre el viernes
            {
                if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("ET") && GridSemana.Rows[fila].Cells["Viernes"].Value != null)
                {
                    ET = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Viernes"].Value.ToString());
                    MessageBox.Show("Estoy en fila ET: " + ET.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Viernes"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("SC") && GridSemana.Rows[fila].Cells["Viernes"].Value != null)
                {
                    SC = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Viernes"].Value.ToString());
                    MessageBox.Show("Estoy en fila SC: " + SC.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Viernes"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("EC") && GridSemana.Rows[fila].Cells["Viernes"].Value != null)
                {
                    EC = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Viernes"].Value.ToString());
                    MessageBox.Show("Estoy en fila EC: " + EC.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Viernes"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("ST") && GridSemana.Rows[fila].Cells["Viernes"].Value != null)
                {
                    ST = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Viernes"].Value.ToString());
                    MessageBox.Show("Estoy en fila ST: " + ST.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Viernes"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("EI") && GridSemana.Rows[fila].Cells["Viernes"].Value != null)
                {
                    EI = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Viernes"].Value.ToString());
                    MessageBox.Show("Estoy en fila EI: " + EI.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Viernes"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("SI") && GridSemana.Rows[fila].Cells["Viernes"].Value != null)
                {
                    SI = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Viernes"].Value.ToString());
                    MessageBox.Show("Estoy en fila SI: " + SI.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Viernes"].Value.ToString()));
                }
            }//For Viernes

            Total = Convert.ToDateTime(Horas[1].ToString()) - Convert.ToDateTime(Horas[0].ToString());
            Total = Total + (Convert.ToDateTime(Horas[3].ToString()) - Convert.ToDateTime(Horas[2].ToString()));
            // Difference in days.
            int Viernes = Total.Hours;
            MessageBox.Show(" Son: " + Viernes.ToString() + " en Viernes");
            /*::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
              -:::::::::::::::::::::::::::::::: SÁBADO :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::*/
            for (int fila = 0; fila < 6; fila++)//Recorre el Sábado
            {
                if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("ET") && GridSemana.Rows[fila].Cells["Sábado"].Value != null)
                {
                    ET = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Sábado"].Value.ToString());
                    MessageBox.Show("Estoy en fila ET: " + ET.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Sábado"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("SC") && GridSemana.Rows[fila].Cells["Sábado"].Value != null)
                {
                    SC = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Sábado"].Value.ToString());
                    MessageBox.Show("Estoy en fila SC: " + SC.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Viernes"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("EC") && GridSemana.Rows[fila].Cells["Sábado"].Value != null)
                {
                    EC = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Sábado"].Value.ToString());
                    MessageBox.Show("Estoy en fila EC: " + EC.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Sábado"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("ST") && GridSemana.Rows[fila].Cells["Sábado"].Value != null)
                {
                    ST = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Sábado"].Value.ToString());
                    MessageBox.Show("Estoy en fila ST: " + ST.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Sábado"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("EI") && GridSemana.Rows[fila].Cells["Sábado"].Value != null)
                {
                    EI = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Sábado"].Value.ToString());
                    MessageBox.Show("Estoy en fila EI: " + EI.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Sábado"].Value.ToString()));
                }
                else if (GridSemana.Rows[fila].Cells["Tipo"].Value.Equals("SI") && GridSemana.Rows[fila].Cells["Sábado"].Value != null)
                {
                    SI = Convert.ToDateTime(GridSemana.Rows[fila].Cells["Sábado"].Value.ToString());
                    MessageBox.Show("Estoy en fila SI: " + SI.ToString("hh:mm"));
                    Horas.Add(Convert.ToDateTime(GridSemana.Rows[fila].Cells["Sábado"].Value.ToString()));
                }
            }//For Sábado

            Total = Convert.ToDateTime(Horas[1].ToString()) - Convert.ToDateTime(Horas[0].ToString());
            Total = Total + (Convert.ToDateTime(Horas[3].ToString()) - Convert.ToDateTime(Horas[2].ToString()));
            // Difference in days.
            int Sábado = Total.Hours;
            MessageBox.Show(" Son: " + Sábado.ToString() + " en Sábado");

            GridSemana.Rows.Add("-", "", "-", "-", "-", "-", "-", "-", "-");
            GridSemana.Rows.Add("", "", "Total Lunes", "Total Martes", "Total Miércoles", "Total Jueves", "Total Viernes", "Total Sábado", "Horas Extras");
            GridSemana.Rows.Add("", "", Lunes, Martes, Miércoles, Jueves, Viernes, Sábado, 0);
        }//Fin de FUNCIÓN

        private void MostrarSemanaPrincipal()
        {
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            ////////////////////////////////// CONSULTA SEMANA DE LOS TRABAJADORES /////////////////////////////////////////////////////////
            DateTime FechaInicio, FechaFinal;
            String FechaFormateada, FinalFormateada;

            FechaInicio = dateInicio.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro
                                           // FechaInicio=FechaInicio.AddDays(-1);//Se agrega 1 dia más
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
                SqlCommand cmd = new SqlCommand("FechaNomina", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FechaInicio", FechaFormateada);
                cmd.Parameters.AddWithValue("@FechaFinal", FinalFormateada);
                cmd.Parameters.AddWithValue("@Planta", listPlantas.Text.ToString());
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(cmd.ExecuteReader());
                GridSemana.DataSource = dt;
                cn.Close();
            }
            catch (Exception ex) { }
            this.ColorReporte();
        }

        public void Columnas(DataGridView grid)//Agrega las columnas manualmente
        {
            if (grid.Name == "GridSemanaHorarios")
            {
                grid.Columns.Add("Trabajador", "Trabajador");
                grid.Columns.Add("Estatus", "Estatus");
                grid.Columns.Add("Lunes", "Lunes");
                grid.Columns.Add("Martes", "Martes");
                grid.Columns.Add("Miércoles", "Miércoles");
                grid.Columns.Add("Jueves", "Jueves");
                grid.Columns.Add("Viernes", "Viernes");
                grid.Columns.Add("Sábado", "Sábado");
            }
            else if (grid.Name == "gridIncidencias")
            {
                grid.Columns.Add("Trabajador", "Trabajador");
                grid.Columns.Add("Turno", "Turno");
                grid.Columns.Add("Estatus", "Estatus");
                grid.Columns.Add("Lunes", "Lunes");
                grid.Columns.Add("Martes", "Martes");
                grid.Columns.Add("Miércoles", "Miércoles");
                grid.Columns.Add("Jueves", "Jueves");
                grid.Columns.Add("Viernes", "Viernes");
                grid.Columns.Add("Sábado", "Sábado");
                //grid.Columns.Add("HorasT", "Horas Totales");
            }
            else
            {
                grid.Columns.Add("Trabajador", "Trabajador");
                grid.Columns.Add("Turno", "Turno");
                grid.Columns.Add("Estatus", "Estatus");
                grid.Columns.Add("Lunes", "Lunes");
                grid.Columns.Add("Martes", "Martes");
                grid.Columns.Add("Miércoles", "Miércoles");
                grid.Columns.Add("Jueves", "Jueves");
                grid.Columns.Add("Viernes", "Viernes");
                grid.Columns.Add("Sábado", "Sábado");
                //grid.Columns.Add("HorasT", "Horas Totales");
            }
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

        private void bARtOP_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
            { xClick = e.X; yClick = e.Y; }
            else
            { this.Left = this.Left + (e.X - xClick); this.Top = this.Top + (e.Y - yClick); }
        }

        public void Filtrar(String Turno)
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
            query = "SELECT DISTINCT Trabajador FROM Checadas WHERE Turno = '" + Turno + "' Order by Trabajador ";

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
                    //no se la ibas a presentar a miros?
                    //:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
                    for (int registro = 0; registro < dt.Rows.Count; registro++)//FOR[4]: Recorre los registros
                    {
                        //   MessageBox.Show("Numero de checada: " + registro);
                        String turnoActual = dt.Rows[registro]["Turno"].ToString();
                        String Estatus = Op.VerEstatus(Empleado[NumNomina].ToString());
                        switch (turnoActual)
                        {
                            case "NOCHE":
                                if (Lunes.Count == 0 || Martes.Count == 0 || Miercoles.Count == 0 || Jueves.Count == 0 || Viernes.Count == 0 || Sabado.Count == 0)
                                {
                                    c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                }
                                else if ((Lunes.Count == 1 || Lunes.Count == NOCHE || Lunes.Count == NOCHE + 2) && (Martes.Count == 1 || Martes.Count == 2 || Martes.Count == NOCHE + 2) && (Miercoles.Count == 1 || Miercoles.Count == 2 || Miercoles.Count == NOCHE + 2) && (Jueves.Count == 1 || Jueves.Count == 2 || Jueves.Count == NOCHE + 2) && (Viernes.Count == 1 || Viernes.Count == NOCHE || Viernes.Count == NOCHE + 2) && (Sabado.Count == 1 || Sabado.Count == 2 || Sabado.Count == NOCHE + 2))//Confirmar que no falte alguna
                                {
                                    if ((Lunes.Count == 1) || (Lunes.Count == (NOCHE + 2)) || (Lunes.Count == NOCHE) && !Estatus.Equals("Error"))//Verificar Todas las checadas del turno y con 2 o 1 checadas extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//LUNES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if ((Martes.Count == 1) || (Martes.Count == (NOCHE + 2)) || (Martes.Count == NOCHE) && !Estatus.Equals("Error"))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//MARTES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if ((Miercoles.Count == 1) || Miercoles.Count == (NOCHE + 2) || (Miercoles.Count == NOCHE) && !Estatus.Equals("Error"))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//MIERCOLES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if ((Jueves.Count == 1) || (Jueves.Count == (NOCHE + 2)) || (Jueves.Count == NOCHE) && !Estatus.Equals("Error"))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//JUEVES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if ((Viernes.Count == 1) || (Viernes.Count == (NOCHE + 2)) || (Viernes.Count == NOCHE) && !Estatus.Equals("Error"))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//VIERNES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if ((Sabado.Count == 1) || (Sabado.Count == (NOCHE + 2)) || (Sabado.Count == NOCHE) && !Estatus.Equals("Error"))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana              
                                    }
                                    else//SABADO
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                }//IF PARES EN TODA LA SEMANA
                                else
                                {
                                    c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                }
                                break;
                            case "SAB":
                                if (Lunes.Count == 0 || Martes.Count == 0 || Miercoles.Count == 0 || Jueves.Count == 0 || Viernes.Count == 0 || Sabado.Count == 0)
                                {
                                    c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                }
                                else if ((Lunes.Count % 2 == 0 || Lunes.Count == 3) && (Martes.Count % 2 == 0 || Martes.Count == 3) && (Miercoles.Count % 2 == 0 || Miercoles.Count == 3) && (Jueves.Count % 2 == 0 || Jueves.Count == 3 || Jueves.Count == 0) && (Viernes.Count % 2 == 0 || Viernes.Count == 3 || Viernes.Count == 0) && (Sabado.Count % 2 == 0 || Sabado.Count == 3))//Confirmar que no falte alguna
                                {
                                    if ((Lunes.Count == (SAB + 2)) || (Lunes.Count == SAB) || (Lunes.Count == (SAB + 1)))//Verificar Todas las checadas del turno y con 2 o 1 checadas extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//LUNES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if ((Martes.Count == (SAB + 2)) || (Martes.Count == SAB) || (Martes.Count == (SAB + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//MARTES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if (Miercoles.Count == (SAB + 2) || Miercoles.Count == SAB || (Miercoles.Count == (SAB + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//MIERCOLES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if (Jueves.Count == (SAB + 2) || Jueves.Count == SAB || (Jueves.Count == (SAB + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//JUEVES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if (Viernes.Count == (SAB + 2) || Viernes.Count == SAB || (Viernes.Count == (SAB + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//VIERNES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if (Sabado.Count == (SAB + 2) || Sabado.Count == SAB || (Sabado.Count == (SAB + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana              
                                    }
                                    else//SABADO
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                }//IF PARES EN TODA LA SEMANA
                                else
                                {
                                    c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                }
                                break;
                            case "TARDE":
                                if (Lunes.Count == 0 || Martes.Count == 0 || Miercoles.Count == 0 || Jueves.Count == 0 || Viernes.Count == 0 || Sabado.Count == 0)
                                {
                                    c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                }
                                else if ((Lunes.Count == 2 || Lunes.Count == 3) && (Martes.Count == 2 || Martes.Count == 3) && (Miercoles.Count == 2 || Miercoles.Count == 3) && (Jueves.Count == 2 || Jueves.Count == 3) && (Viernes.Count == 2 || Viernes.Count == 3) && (Sabado.Count == 2 || Sabado.Count == 3))//Confirmar que no falte alguna
                                {
                                    if ((Lunes.Count == TARDE) || (Lunes.Count == (TARDE + 1)))//Verificar Todas las checadas del turno y con 2 o 1 checadas extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//LUNES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if ((Martes.Count == TARDE) || (Martes.Count == (TARDE + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//MARTES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if (Miercoles.Count == TARDE || (Miercoles.Count == (TARDE + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//MIERCOLES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if (Jueves.Count == TARDE || (Jueves.Count == (TARDE + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//JUEVES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if (Viernes.Count == TARDE || (Viernes.Count == (TARDE + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//VIERNES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if (Sabado.Count == TARDE || (Sabado.Count == (TARDE + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana              
                                    }
                                    else//SABADO
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                }//IF PARES EN TODA LA SEMANA
                                else
                                {
                                    c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                }
                                break;
                            case "TDIA":
                                if (Lunes.Count == 0 || Lunes.Count < 2 || Martes.Count == 0 || Martes.Count < 2 || Miercoles.Count == 0 || Miercoles.Count < 2 || Jueves.Count == 0 || Jueves.Count < 2 || Viernes.Count == 0 || Viernes.Count < 2 || Sabado.Count == 0 || Sabado.Count < 2)
                                {
                                    c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                }
                                else if ((Lunes.Count == 6 || Lunes.Count == 5 || Lunes.Count == TDIA) && (Martes.Count == 6 || Martes.Count == 5 || Martes.Count == TDIA) && (Miercoles.Count == 6 || Miercoles.Count == 5 || Miercoles.Count == TDIA) && (Jueves.Count == 6 || Jueves.Count == 5 || Jueves.Count == TDIA) && (Viernes.Count == 6 || Viernes.Count == 5 || Viernes.Count == TDIA) && (Sabado.Count == 6 || Sabado.Count == 5 || Sabado.Count == TDIA))//Confirmar que no falte alguna
                                {
                                    if ((Lunes.Count == TDIA) || (Lunes.Count == (TDIA + 2)) || (Lunes.Count == (TDIA + 1)))//Verificar Todas las checadas del turno y con 2 o 1 checadas extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//LUNES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }

                                    if ((Martes.Count == (TDIA + 2)) || (Martes.Count == TDIA) || (Martes.Count == (TDIA + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//MARTES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if (Miercoles.Count == (TDIA + 2) || Miercoles.Count == TDIA || (Miercoles.Count == (TDIA + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//MIERCOLES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if (Jueves.Count == (TDIA + 2) || Jueves.Count == TDIA || (Jueves.Count == (TDIA + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//JUEVES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if (Viernes.Count == (TDIA + 2) || Viernes.Count == TDIA || (Viernes.Count == (TDIA + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//VIERNES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if (Sabado.Count == (TDIA + 2) || Sabado.Count == TDIA || (Sabado.Count == (TDIA + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana              
                                    }
                                    else//SABADO
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                }//IF PARES EN TODA LA SEMANA
                                else
                                {
                                    c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                }
                                break;
                            case "TFIJ":
                                if (Lunes.Count == 0 || Martes.Count == 0 || Miercoles.Count == 0 || Jueves.Count == 0 || Viernes.Count == 0 || Sabado.Count == 0)
                                {
                                    c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                }
                                else if ((Lunes.Count == 6 || Lunes.Count == 5 || Lunes.Count == TFIJ) && (Martes.Count == 6 || Martes.Count == 5 || Martes.Count == TFIJ) && (Miercoles.Count == 6 || Miercoles.Count == 5 || Miercoles.Count == TFIJ) && (Jueves.Count == 6 || Jueves.Count == 5 || Jueves.Count == TFIJ) && (Viernes.Count == 6 || Viernes.Count == 5 || Viernes.Count == TFIJ) && (Sabado.Count == 3 || Sabado.Count == 2))//Confirmar que no falte alguna
                                {
                                    if ((Lunes.Count == (TFIJ + 2)) || (Lunes.Count == TFIJ) || (Lunes.Count == (TFIJ + 1)))//Verificar Todas las checadas del turno y con 2 o 1 checadas extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//LUNES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if ((Martes.Count == (TFIJ + 2)) || (Martes.Count == TFIJ) || (Martes.Count == (TFIJ + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//MARTES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if (Miercoles.Count == (TFIJ + 2) || Miercoles.Count == TFIJ || (Miercoles.Count == (TFIJ + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//MIERCOLES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if (Jueves.Count == (TFIJ + 2) || Jueves.Count == TFIJ || (Jueves.Count == (TFIJ + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//JUEVES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if (Viernes.Count == (TFIJ + 2) || Viernes.Count == TFIJ || (Viernes.Count == (TFIJ + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//VIERNES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if ((Sabado.Count == 2) || Sabado.Count == (TFIJ + 2) || (Sabado.Count == TFIJ) || (Sabado.Count == (TFIJ + 1)))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana              
                                    }
                                    else//SABADO
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                }//IF PARES EN TODA LA SEMANA
                                else
                                {
                                    c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                }
                                break;
                            case "TNO2":
                                if (Lunes.Count == 0 || Martes.Count == 0 || Miercoles.Count == 0 || Jueves.Count == 0 || Viernes.Count == 0 || Sabado.Count == 0)
                                {
                                    c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                }
                                else if ((Lunes.Count % 2 == 0) && (Martes.Count % 2 == 0) && (Miercoles.Count % 2 == 0) && (Jueves.Count % 2 == 0) && (Viernes.Count % 2 == 0) && (Sabado.Count % 2 == 0))//Confirmar que no falte alguna
                                {
                                    if ((Lunes.Count == (TNO2 + 2)) || (Lunes.Count == TNO2))//Verificar Todas las checadas del turno y con 2 o 1 checadas extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//LUNES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if ((Martes.Count == (TNO2 + 2)) || (Martes.Count == TNO2))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//MARTES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if (Miercoles.Count == (TNO2 + 2) || (Miercoles.Count == TNO2))//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//MIERCOLES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if (Jueves.Count == (TNO2 + 2) || Jueves.Count == TNO2)//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//JUEVES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if (Viernes.Count == (TNO2 + 2) || Viernes.Count == TNO2)//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana
                                    }
                                    else//VIERNES
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                    if (Sabado.Count == (TNO2 + 2) || Sabado.Count == TNO2)//Verificar Todas las checadas del turno o con extras
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Correcto");//Actualizo el estatus de la semana              
                                    }
                                    else//SABADO
                                    {
                                        c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                    }
                                }//IF PARES EN TODA LA SEMANA
                                else
                                {
                                    c.ActualizarEstatus(Empleado[NumNomina].ToString(), "Error");//Actualizo el estatus de la semana
                                }
                                break;
                            default:
                                break;
                        }//SWITCH
                    }//FOR [4]: Recorre Registros DT 
                }
                catch (Exception ex) { MessageBox.Show("No entro al try el Empleado" + Empleado[NumNomina].ToString() + ex); }
            }//FOR[2] RECORRE TRABAJADORES
        }//Fin de funcion Filtrado

        private void Descontar_Click(object sender, EventArgs e)//Quitar tiempo por retardos
        {
            if (GridSemana.Rows.Count > 1)
            {
                for (int i = GridSemana.Rows.Count - 1; i >= 0; i--)
                {
                    GridSemana.Rows.RemoveAt(i);
                }
            }
            GridSemana.Refresh();
            RH.LimpiarListaRetardos();
            DateTime EntradaNoche = Convert.ToDateTime("19:00:00");
            DateTime EntradaSAB = Convert.ToDateTime("08:00:00");
            DateTime EntradaTARDE = Convert.ToDateTime("15:30:00");
            DateTime EntradaTDIA = Convert.ToDateTime("07:00:00");
            DateTime EntradaTFIJ = Convert.ToDateTime("08:00:00");
            DateTime EntradaTNO2 = Convert.ToDateTime("07:00:00");

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
            string query = "SELECT DISTINCT Trabajador FROM Checadas Order by Trabajador ";

            SqlCommand cmd = new SqlCommand(query, cn);
            DataTable t1 = new DataTable();
            using (SqlDataAdapter a = new SqlDataAdapter(cmd))
            {
                a.Fill(t1);
            }

            for (int x = 0; x < t1.Rows.Count; x++)//FOR[1] : Agrega los Trabajadores a una lista
            {
                Empleado.Add(t1.Rows[x]["Trabajador"].ToString());
            }

            BarraProgreso.Maximum = Empleado.Count;
            BarraProgreso.BackColor = Color.White;
            BarraProgreso.ProgressColor1 = Color.CornflowerBlue;
            BarraProgreso.ProgressColor2 = Color.DarkRed;
            BarraProgreso.Visible = true;
            BarraProgreso.ForeColor = Color.Cyan;
            ///////////////////////////////////////////// OBTENER DATOS DE CADA EMPLEADO EN LISTA //////////////////////////////////////////////////////
            ///
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            for (int indiceEmpleados = 0; indiceEmpleados < Empleado.Count; indiceEmpleados++)
            {
                BarraProgreso.Value = indiceEmpleados;
                BarraProgreso.Update();//Contador de progreso

                Lunes.Clear();
                Martes.Clear();
                Miercoles.Clear();
                Jueves.Clear();
                Viernes.Clear();
                Viernes.Clear();
                Sabado.Clear();

                query = "SELECT * FROM Checadas WHERE Trabajador = '" + Empleado[indiceEmpleados] + "' Order by Trabajador,Checada ";

                SqlCommand cmd2 = new SqlCommand(query, cn);
                DataTable dt = new DataTable();
                using (SqlDataAdapter a = new SqlDataAdapter(cmd2))
                {
                    a.Fill(dt);
                }

                /////////////////////////////////////////////// OBTENER CHECADAS DEL DIA //////////////////////////////////////////////////////
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

                /////////////////////////////////////////////// REVISAR ENTRADA DIARIA //////////////////////////////////////////////////////

                for (int indiceTurno = 0; indiceTurno < dt.Rows.Count; indiceTurno++)//Recorrer cada registro para obtener el turno
                {
                    ////////////////////////////////// Codigo para el dia de la semana /////////////////////////////////////////////////////
                    DateTime dateValue = new DateTime();
                    CultureInfo cul = CultureInfo.CurrentCulture;
                    // Usa la fecha formateada y calcula el número de la semana
                    int NumeroSemana = cul.Calendar.GetWeekOfYear(dateValue, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
                    dateValue = Convert.ToDateTime(dt.Rows[indiceTurno]["Checada"].ToString());//Consulto la checada del indice
                    String dia = dateValue.DayOfWeek.ToString();//Convierto la fecha en dia de la semana
                    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                    String TurnoRegistro = dt.Rows[indiceTurno]["Turno"].ToString();
                    switch (dia)
                    {
                        case "Lunes":
                        case "Monday":
                            try
                            {
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE"))//TURNO NOCHE
                                {
                                    TimeSpan Noche = new TimeSpan(19, 00, 0);//Variable con el horario 19:00 hrs
                                    EntradaNoche = Convert.ToDateTime(Lunes[0].ToString()).Date + Noche;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Lunes ----------------------------------------------------------------------//
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Lunes[0].ToString()) != null) && (Convert.ToDateTime(Lunes[0].ToString()) < EntradaNoche.AddMinutes(-30)))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaNoche);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Lunes[0].ToString()) != null) && Convert.ToDateTime(Lunes[0].ToString()) >= (EntradaNoche.AddMinutes(-30)) && Convert.ToDateTime(Lunes[0].ToString()) < (EntradaNoche.AddMinutes(6)))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaNoche);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Lunes[0].ToString()) != null) && Convert.ToDateTime(Lunes[0].ToString()) > (EntradaNoche.AddMinutes(6)))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaNoche);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Lunes[0].ToString(), "NOCHE");
                                    }
                                }//fin noche
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB"))//TURNO
                                {
                                    TimeSpan SAB = new TimeSpan(08, 00, 0);//Variable con el horario 08:00 hrs
                                    EntradaSAB = Convert.ToDateTime(Lunes[0].ToString()).Date + SAB;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Lunes ----------------------------------------------------------------------//
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Lunes[0].ToString()) != null) && (Convert.ToDateTime(Lunes[0].ToString()) < EntradaSAB.AddMinutes(-30)))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaSAB);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Lunes[0].ToString()) != null) && Convert.ToDateTime(Lunes[0].ToString()) >= (EntradaSAB.AddMinutes(-30)) && Convert.ToDateTime(Lunes[0].ToString()) < (EntradaSAB.AddMinutes(6)))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaSAB);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Lunes[0].ToString()) != null) && Convert.ToDateTime(Lunes[0].ToString()) > (EntradaSAB.AddMinutes(6)))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaSAB);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Lunes[0].ToString(), "SAB");
                                    }
                                }//fin sab
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE"))//TURNO
                                {
                                    TimeSpan TARDE = new TimeSpan(15, 30, 0);//Variable con el horario 15:30 hrs
                                    EntradaTARDE = Convert.ToDateTime(Lunes[0].ToString()).Date + TARDE;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Lunes ----------------------------------------------------------------------//
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Lunes[0].ToString()) != null) && (Convert.ToDateTime(Lunes[0].ToString()) < EntradaTARDE.AddMinutes(-30)))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaTARDE);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Lunes[0].ToString()) != null) && (Convert.ToDateTime(Lunes[0].ToString()) >= (EntradaTARDE.AddMinutes(-30))) && (Convert.ToDateTime(Lunes[0].ToString()) < (EntradaTARDE.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaTARDE);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Lunes[0].ToString()) != null) && Convert.ToDateTime(Lunes[0].ToString()) > (EntradaTARDE.AddMinutes(6)))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaTARDE);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Lunes[0].ToString(), "TARDE");
                                    }
                                }//fin tarde
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA"))//TURNO
                                {
                                    TimeSpan TDIA = new TimeSpan(07, 00, 0);//Variable con el horario 07:00 hrs
                                    EntradaTDIA = Convert.ToDateTime(Lunes[0].ToString()).Date + TDIA;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Lunes ----------------------------------------------------------------------//
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Lunes[0].ToString()) != null) && (Convert.ToDateTime(Lunes[0].ToString()) < EntradaTDIA.AddMinutes(-30)))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaTDIA);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Lunes[0].ToString()) != null) && (Convert.ToDateTime(Lunes[0].ToString()) >= (EntradaTDIA.AddMinutes(-30))) && (Convert.ToDateTime(Lunes[0].ToString()) < (EntradaTDIA.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaTDIA);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Lunes[0].ToString()) != null) && (Convert.ToDateTime(Lunes[0].ToString()) > (EntradaTDIA.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaTDIA);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Lunes[0].ToString(), "TDIA");
                                    }
                                }//tdia
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ"))//TURNO
                                {
                                    TimeSpan TFIJ = new TimeSpan(08, 00, 0);//Variable con el horario 07:00 hrs
                                    EntradaTFIJ = Convert.ToDateTime(Lunes[0].ToString()).Date + TFIJ;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Lunes ----------------------------------------------------------------------//
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Lunes[0].ToString()) != null) && (Convert.ToDateTime(Lunes[0].ToString()) < EntradaTFIJ.AddMinutes(-30)))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaTFIJ);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Lunes[0].ToString()) != null) && (Convert.ToDateTime(Lunes[0].ToString()) >= (EntradaTFIJ.AddMinutes(-30))) && (Convert.ToDateTime(Lunes[0].ToString()) < (EntradaTFIJ.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaTFIJ);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Lunes[0].ToString()) != null) && (Convert.ToDateTime(Lunes[0].ToString()) > (EntradaTFIJ.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaTFIJ);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Lunes[0].ToString(), "TFIJ");
                                    }
                                }//tfij
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2"))//TURNO
                                {
                                    TimeSpan TNO2 = new TimeSpan(07, 00, 0);//Variable con el horario 07:00 hrs
                                    EntradaTNO2 = Convert.ToDateTime(Lunes[0].ToString()).Date + TNO2;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Lunes ----------------------------------------------------------------------//
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Lunes[0].ToString()) != null) && (Convert.ToDateTime(Lunes[0].ToString()) < EntradaTNO2.AddMinutes(-30)))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaTNO2);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Lunes[0].ToString()) != null) && (Convert.ToDateTime(Lunes[0].ToString()) >= (EntradaTNO2.AddMinutes(-30))) && (Convert.ToDateTime(Lunes[0].ToString()) < (EntradaTNO2.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaTNO2);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Lunes[0].ToString()) != null) && (Convert.ToDateTime(Lunes[0].ToString()) > (EntradaTNO2.AddMinutes(6))))
                                    {
                                        //MessageBox.Show("Entra al if con la checada: " + Lunes[0].ToString() + " Y Entrada es de: " + EntradaTNO2.ToString());
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), EntradaTNO2);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Lunes[0].ToString(), "TNO2");
                                    }
                                }//tno2
                            } catch (Exception ex) { }//CATCH de Dia Lunes
                            break;

                        case "Martes":
                        case "Tuesday":
                            try
                            {
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE"))//TURNO NOCHE
                                {
                                    TimeSpan Noche = new TimeSpan(19, 00, 0);//Variable con el horario 19:00 hrs
                                                                             //----------------------------------------------------------- Martes ----------------------------------------------------------------------//
                                    EntradaNoche = Convert.ToDateTime(Martes[1].ToString()).Date + Noche;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Martes[1].ToString()) != null) && (Convert.ToDateTime(Martes[1].ToString()) < (EntradaNoche.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[1].ToString()), EntradaNoche);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Martes[1].ToString()) != null) && (Convert.ToDateTime(Martes[1].ToString()) >= (EntradaNoche.AddMinutes(-30))) && (Convert.ToDateTime(Martes[1].ToString()) < (EntradaNoche.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[1].ToString()), EntradaNoche);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Martes[1].ToString()) != null) && (Convert.ToDateTime(Martes[1].ToString()) > (EntradaNoche.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[1].ToString()), EntradaNoche);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Martes[1].ToString(), "NOCHE");
                                    }
                                }//fin noche
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB"))//TURNO SAB
                                {
                                    TimeSpan SAB = new TimeSpan(08, 00, 0);//Variable con el horario 08:00 hrs
                                    //----------------------------------------------------------- Martes ----------------------------------------------------------------------//
                                    EntradaSAB = Convert.ToDateTime(Martes[1].ToString()).Date + SAB;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Martes[0].ToString()) != null) && (Convert.ToDateTime(Martes[0].ToString()) < (EntradaSAB.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), EntradaSAB);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Martes[0].ToString()) != null) && (Convert.ToDateTime(Martes[0].ToString()) >= (EntradaSAB.AddMinutes(-30))) && (Convert.ToDateTime(Martes[0].ToString()) < (EntradaSAB.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), EntradaSAB);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Martes[0].ToString()) != null) && (Convert.ToDateTime(Martes[0].ToString()) > (EntradaSAB.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), EntradaSAB);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Martes[0].ToString(), "SAB");
                                    }
                                }//fin sab
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE"))//TURNO TARDE
                                {
                                    TimeSpan TARDE = new TimeSpan(15, 30, 0);//Variable con el horario 15:30 hrs
                                    //----------------------------------------------------------- Martes ----------------------------------------------------------------------//
                                    EntradaTARDE = Convert.ToDateTime(Martes[0].ToString()).Date + TARDE;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Martes[0].ToString()) != null) && (Convert.ToDateTime(Martes[0].ToString()) < (EntradaTARDE.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), EntradaTARDE);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Martes[0].ToString()) != null) && (Convert.ToDateTime(Martes[0].ToString()) >= (EntradaTARDE.AddMinutes(-30))) && (Convert.ToDateTime(Martes[0].ToString()) < (EntradaTARDE.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), EntradaTARDE);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Martes[0].ToString()) != null) && (Convert.ToDateTime(Martes[0].ToString()) > (EntradaTARDE.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), EntradaTARDE);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Martes[0].ToString(), "TARDE");
                                    }
                                }//fin tarde
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA"))//TURNO TDIA
                                {
                                    TimeSpan TDIA = new TimeSpan(07, 00, 0);//Variable con el horario 07:00 hrs
                                    //----------------------------------------------------------- Martes ----------------------------------------------------------------------//
                                    EntradaTDIA = Convert.ToDateTime(Martes[0].ToString()).Date + TDIA;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Martes[0].ToString()) != null) && (Convert.ToDateTime(Martes[0].ToString()) < (EntradaTDIA.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), EntradaTDIA);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Martes[0].ToString()) != null) && (Convert.ToDateTime(Martes[0].ToString()) >= (EntradaTDIA.AddMinutes(-30))) && (Convert.ToDateTime(Martes[0].ToString()) < (EntradaTDIA.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), EntradaTDIA);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Martes[0].ToString()) != null) && (Convert.ToDateTime(Martes[0].ToString()) > (EntradaTDIA.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), EntradaTDIA);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Martes[0].ToString(), "TDIA");
                                    }
                                }//fin tdia
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ"))//TURNO TFIJ
                                {
                                    TimeSpan TFIJ = new TimeSpan(08, 00, 0);//Variable con el horario 07:00 hrs
                                    //----------------------------------------------------------- Martes ----------------------------------------------------------------------//
                                    EntradaTFIJ = Convert.ToDateTime(Martes[0].ToString()).Date + TFIJ;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Martes[0].ToString()) != null) && (Convert.ToDateTime(Martes[0].ToString()) < (EntradaTFIJ.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), EntradaTFIJ);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Martes[0].ToString()) != null) && (Convert.ToDateTime(Martes[0].ToString()) >= (EntradaTFIJ.AddMinutes(-30))) && (Convert.ToDateTime(Martes[0].ToString()) < (EntradaTFIJ.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), EntradaTFIJ);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Martes[0].ToString()) != null) && (Convert.ToDateTime(Martes[0].ToString()) > (EntradaTFIJ.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), EntradaTFIJ);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Martes[0].ToString(), "TFIJ");
                                    }
                                }//fin tfij
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2"))//TURNO TNO2
                                {
                                    TimeSpan TNO2 = new TimeSpan(07, 00, 0);//Variable con el horario 07:00 hrs

                                    //----------------------------------------------------------- Martes ----------------------------------------------------------------------//
                                    EntradaTNO2 = Convert.ToDateTime(Martes[0].ToString()).Date + TNO2;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Martes[0].ToString()) != null) && (Convert.ToDateTime(Martes[0].ToString()) < (EntradaTNO2.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), EntradaTNO2);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Martes[0].ToString()) != null) && (Convert.ToDateTime(Martes[0].ToString()) >= (EntradaTNO2.AddMinutes(-30))) && (Convert.ToDateTime(Martes[0].ToString()) < (EntradaTNO2.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), EntradaTNO2);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Martes[0].ToString()) != null) && (Convert.ToDateTime(Martes[0].ToString()) > (EntradaTNO2.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), EntradaTNO2);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Martes[0].ToString(), "TNO2");
                                    }
                                }//fin tno2
                            } catch (Exception ex) { }//fin martes
                            break;

                        case "Miércoles":
                        case "Wednesday":
                            try
                            {
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE"))//TURNO NOCHE
                                {
                                    TimeSpan Noche = new TimeSpan(19, 00, 0);//Variable con el horario 19:00 hrs
                                                                             //----------------------------------------------------------- Miercoles ----------------------------------------------------------------------//
                                    EntradaNoche = Convert.ToDateTime(Miercoles[1].ToString()).Date + Noche;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Miercoles[1].ToString()) != null) && (Convert.ToDateTime(Miercoles[1].ToString()) < (EntradaNoche.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[1].ToString()), EntradaNoche);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Miercoles[1].ToString()) != null) && (Convert.ToDateTime(Miercoles[1].ToString()) >= (EntradaNoche.AddMinutes(-30))) && (Convert.ToDateTime(Miercoles[1].ToString()) < (EntradaNoche.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[1].ToString()), EntradaNoche);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Miercoles[1].ToString()) != null) && (Convert.ToDateTime(Miercoles[1].ToString()) > (EntradaNoche.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[1].ToString()), EntradaNoche);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Miercoles[1].ToString(), "NOCHE");
                                    }
                                }//fin noche
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB"))//TURNO SAB
                                {
                                    TimeSpan SAB = new TimeSpan(08, 00, 0);//Variable con el horario 08:00 hrs  
                                                                           //----------------------------------------------------------- Miercoles ----------------------------------------------------------------------//
                                    EntradaSAB = Convert.ToDateTime(Miercoles[0].ToString()).Date + SAB;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Miercoles[0].ToString()) != null) && (Convert.ToDateTime(Miercoles[0].ToString()) < (EntradaSAB.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), EntradaSAB);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Miercoles[0].ToString()) != null) && (Convert.ToDateTime(Miercoles[0].ToString()) >= (EntradaSAB.AddMinutes(-30))) && (Convert.ToDateTime(Miercoles[0].ToString()) < (EntradaSAB.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), EntradaSAB);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Miercoles[0].ToString()) != null) && (Convert.ToDateTime(Miercoles[0].ToString()) > (EntradaSAB.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), EntradaSAB);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Miercoles[0].ToString(), "SAB");
                                    }
                                }//fin sab
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE"))//TURNO TARDE
                                {
                                    TimeSpan TARDE = new TimeSpan(15, 30, 0);//Variable con el horario 15:30 hrs
                                                                             //----------------------------------------------------------- Miercoles ----------------------------------------------------------------------//
                                    EntradaTARDE = Convert.ToDateTime(Miercoles[0].ToString()).Date + TARDE;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Miercoles[0].ToString()) != null) && (Convert.ToDateTime(Miercoles[0].ToString()) < (EntradaTARDE.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), EntradaTARDE);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Miercoles[0].ToString()) != null) && (Convert.ToDateTime(Miercoles[0].ToString()) >= (EntradaTARDE.AddMinutes(-30))) && (Convert.ToDateTime(Miercoles[0].ToString()) < (EntradaTARDE.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), EntradaTARDE);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Miercoles[0].ToString()) != null) && (Convert.ToDateTime(Miercoles[0].ToString()) > (EntradaTARDE.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), EntradaTARDE);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Miercoles[0].ToString(), "TARDE");
                                    }
                                }//fin tarde
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA"))//TURNO TDIA
                                {
                                    TimeSpan TDIA = new TimeSpan(07, 00, 0);//Variable con el horario 07:00 hrs
                                                                            //----------------------------------------------------------- Miercoles ----------------------------------------------------------------------//
                                    EntradaTDIA = Convert.ToDateTime(Miercoles[0].ToString()).Date + TDIA;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Miercoles[0].ToString()) != null) && (Convert.ToDateTime(Miercoles[0].ToString()) < (EntradaTDIA.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), EntradaTDIA);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Miercoles[0].ToString()) != null) && (Convert.ToDateTime(Miercoles[0].ToString()) >= (EntradaTDIA.AddMinutes(-30))) && (Convert.ToDateTime(Miercoles[0].ToString()) < (EntradaTDIA.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), EntradaTDIA);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Miercoles[0].ToString()) != null) && (Convert.ToDateTime(Miercoles[0].ToString()) > (EntradaTDIA.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), EntradaTDIA);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Miercoles[0].ToString(), "TDIA");
                                    }
                                }//fin tdia
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ"))//TURNO TFIJ
                                {
                                    TimeSpan TFIJ = new TimeSpan(08, 00, 0);//Variable con el horario 07:00 hrs
                                    //----------------------------------------------------------- Miercoles ----------------------------------------------------------------------//
                                    EntradaTFIJ = Convert.ToDateTime(Miercoles[0].ToString()).Date + TFIJ;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Miercoles[0].ToString()) != null) && (Convert.ToDateTime(Miercoles[0].ToString()) < (EntradaTFIJ.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), EntradaTFIJ);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Miercoles[0].ToString()) != null) && (Convert.ToDateTime(Miercoles[0].ToString()) >= (EntradaTFIJ.AddMinutes(-30))) && (Convert.ToDateTime(Miercoles[0].ToString()) < (EntradaTFIJ.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), EntradaTFIJ);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Miercoles[0].ToString()) != null) && (Convert.ToDateTime(Miercoles[0].ToString()) > (EntradaTFIJ.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), EntradaTFIJ);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Miercoles[0].ToString(), "TFIJ");
                                    }
                                }//fin tfij
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2"))//TURNO TNO2
                                {
                                    TimeSpan TNO2 = new TimeSpan(07, 00, 0);//Variable con el horario 07:00 hrs
                                                                            //----------------------------------------------------------- Miercoles ----------------------------------------------------------------------//
                                    EntradaTNO2 = Convert.ToDateTime(Miercoles[0].ToString()).Date + TNO2;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Miercoles[0].ToString()) != null) && (Convert.ToDateTime(Miercoles[0].ToString()) < (EntradaTNO2.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), EntradaTNO2);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Miercoles[0].ToString()) != null) && (Convert.ToDateTime(Miercoles[0].ToString()) >= (EntradaTNO2.AddMinutes(-30))) && (Convert.ToDateTime(Miercoles[0].ToString()) < (EntradaTNO2.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), EntradaTNO2);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Miercoles[0].ToString()) != null) && (Convert.ToDateTime(Miercoles[0].ToString()) > (EntradaTNO2.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), EntradaTNO2);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Miercoles[0].ToString(), "TNO2");
                                    }
                                }//fin tno2
                            } catch (Exception ex) { }
                            break;

                        case "Jueves":
                        case "Thursday":
                            try
                            {
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE"))//TURNO NOCHE
                                {
                                    TimeSpan Noche = new TimeSpan(19, 00, 0);//Variable con el horario 19:00 hrs
                                    //----------------------------------------------------------- Jueves ----------------------------------------------------------------------//
                                    EntradaNoche = Convert.ToDateTime(Jueves[1].ToString()).Date + Noche;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Jueves[1].ToString()) != null) && (Convert.ToDateTime(Jueves[1].ToString()) < (EntradaNoche.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[1].ToString()), EntradaNoche);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Jueves[1].ToString()) != null) && (Convert.ToDateTime(Jueves[1].ToString()) >= (EntradaNoche.AddMinutes(-30))) && (Convert.ToDateTime(Jueves[1].ToString()) < (EntradaNoche.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[1].ToString()), EntradaNoche);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Jueves[1].ToString()) != null) && (Convert.ToDateTime(Jueves[1].ToString()) > (EntradaNoche.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[1].ToString()), EntradaNoche);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Jueves[1].ToString(), "NOCHE");
                                    }
                                }//fin noche
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB"))//TURNO SAB
                                {
                                    TimeSpan SAB = new TimeSpan(08, 00, 0);//Variable con el horario 08:00 hrs 
                                    //----------------------------------------------------------- Jueves ----------------------------------------------------------------------//
                                    EntradaSAB = Convert.ToDateTime(Jueves[0].ToString()).Date + SAB;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Jueves[0].ToString()) != null) && (Convert.ToDateTime(Jueves[0].ToString()) < (EntradaSAB.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), EntradaSAB);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Jueves[0].ToString()) != null) && (Convert.ToDateTime(Jueves[0].ToString()) >= (EntradaSAB.AddMinutes(-30))) && (Convert.ToDateTime(Jueves[0].ToString()) < (EntradaSAB.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), EntradaSAB);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Jueves[0].ToString()) != null) && (Convert.ToDateTime(Jueves[0].ToString()) > (EntradaSAB.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), EntradaSAB);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Jueves[0].ToString(), "SAB");
                                    }
                                }//FIN SAB
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE"))//TURNO TARDE
                                {
                                    TimeSpan TARDE = new TimeSpan(15, 30, 0);//Variable con el horario 15:30 hrs
                                    //----------------------------------------------------------- Jueves ----------------------------------------------------------------------//
                                    EntradaTARDE = Convert.ToDateTime(Jueves[0].ToString()).Date + TARDE;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Jueves[0].ToString()) != null) && (Convert.ToDateTime(Jueves[0].ToString()) < (EntradaTARDE.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), EntradaTARDE);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Jueves[0].ToString()) != null) && (Convert.ToDateTime(Jueves[0].ToString()) >= (EntradaTARDE.AddMinutes(-30))) && (Convert.ToDateTime(Jueves[0].ToString()) < (EntradaTARDE.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), EntradaTARDE);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Jueves[0].ToString()) != null) && (Convert.ToDateTime(Jueves[0].ToString()) > (EntradaTARDE.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), EntradaTARDE);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Jueves[0].ToString(), "TARDE");
                                    }
                                }//FIN TARDE
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA"))//TURNO TDIA
                                {
                                    TimeSpan TDIA = new TimeSpan(07, 00, 0);//Variable con el horario 07:00 hrs
                                    //----------------------------------------------------------- Jueves ----------------------------------------------------------------------//
                                    EntradaTDIA = Convert.ToDateTime(Jueves[0].ToString()).Date + TDIA;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Jueves[0].ToString()) != null) && (Convert.ToDateTime(Jueves[0].ToString()) < (EntradaTDIA.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), EntradaTDIA);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Jueves[0].ToString()) != null) && (Convert.ToDateTime(Jueves[0].ToString()) >= (EntradaTDIA.AddMinutes(-30))) && (Convert.ToDateTime(Jueves[0].ToString()) < (EntradaTDIA.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), EntradaTDIA);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Jueves[0].ToString()) != null) && (Convert.ToDateTime(Jueves[0].ToString()) > (EntradaTDIA.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), EntradaTDIA);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Jueves[0].ToString(), "TDIA");
                                    }
                                }//FIN TDIA
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ"))//TURNO TFIJ
                                {
                                    TimeSpan TFIJ = new TimeSpan(08, 00, 0);//Variable con el horario 07:00 hrs
                                    //----------------------------------------------------------- Jueves ----------------------------------------------------------------------//
                                    EntradaTFIJ = Convert.ToDateTime(Jueves[0].ToString()).Date + TFIJ;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Jueves[0].ToString()) != null) && (Convert.ToDateTime(Jueves[0].ToString()) < (EntradaTFIJ.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), EntradaTFIJ);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Jueves[0].ToString()) != null) && (Convert.ToDateTime(Jueves[0].ToString()) >= (EntradaTFIJ.AddMinutes(-30))) && (Convert.ToDateTime(Jueves[0].ToString()) < (EntradaTFIJ.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), EntradaTFIJ);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Jueves[0].ToString()) != null) && (Convert.ToDateTime(Jueves[0].ToString()) > (EntradaTFIJ.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), EntradaTFIJ);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Jueves[0].ToString(), "TFIJ");
                                    }
                                }//FIN TFIJ
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2"))//TURNO TNO2
                                {
                                    TimeSpan TNO2 = new TimeSpan(07, 00, 0);//Variable con el horario 07:00 hrs
                                    //----------------------------------------------------------- Jueves ----------------------------------------------------------------------//
                                    EntradaTNO2 = Convert.ToDateTime(Jueves[0].ToString()).Date + TNO2;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Jueves[0].ToString()) != null) && (Convert.ToDateTime(Jueves[0].ToString()) < (EntradaTNO2.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), EntradaTNO2);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Jueves[0].ToString()) != null) && (Convert.ToDateTime(Jueves[0].ToString()) >= (EntradaTNO2.AddMinutes(-30))) && (Convert.ToDateTime(Jueves[0].ToString()) < (EntradaTNO2.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), EntradaTNO2);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Jueves[0].ToString()) != null) && (Convert.ToDateTime(Jueves[0].ToString()) > (EntradaTNO2.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), EntradaTNO2);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Jueves[0].ToString(), "TNO2");
                                    }
                                }//FIN TNO2
                            } catch (Exception ex) { }
                            break;

                        case "Viernes":
                        case "Friday":
                            try
                            {
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE"))//TURNO NOCHE
                                {
                                    TimeSpan Noche = new TimeSpan(19, 00, 0);//Variable con el horario 19:00 hrs
                                    //----------------------------------------------------------- Viernes ----------------------------------------------------------------------//
                                    EntradaNoche = Convert.ToDateTime(Viernes[1].ToString()).Date + Noche;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Viernes[1].ToString()) != null) && (Convert.ToDateTime(Viernes[1].ToString()) < (EntradaNoche.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[1].ToString()), EntradaNoche);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Viernes[1].ToString()) != null) && (Convert.ToDateTime(Viernes[1].ToString()) >= (EntradaNoche.AddMinutes(-30))) && (Convert.ToDateTime(Viernes[1].ToString()) < (EntradaNoche.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[1].ToString()), EntradaNoche);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Viernes[1].ToString()) != null) && (Convert.ToDateTime(Viernes[1].ToString()) > (EntradaNoche.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[1].ToString()), EntradaNoche);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Viernes[1].ToString(), "NOCHE");
                                    }
                                }//fin noche
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB"))//TURNO SAB
                                {
                                    TimeSpan SAB = new TimeSpan(08, 00, 0);//Variable con el horario 08:00 hrs 
                                    //----------------------------------------------------------- Viernes ----------------------------------------------------------------------//
                                    EntradaSAB = Convert.ToDateTime(Viernes[0].ToString()).Date + SAB;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Viernes[0].ToString()) != null) && (Convert.ToDateTime(Viernes[0].ToString()) < (EntradaSAB.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), EntradaSAB);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Viernes[0].ToString()) != null) && (Convert.ToDateTime(Viernes[0].ToString()) >= (EntradaSAB.AddMinutes(-30))) && (Convert.ToDateTime(Viernes[0].ToString()) < (EntradaSAB.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), EntradaSAB);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Viernes[0].ToString()) != null) && (Convert.ToDateTime(Viernes[0].ToString()) > (EntradaSAB.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), EntradaSAB);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Viernes[0].ToString(), "SAB");
                                    }
                                }
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE"))//TURNO SAB
                                {
                                    TimeSpan TARDE = new TimeSpan(15, 30, 0);//Variable con el horario 15:30 hrs
                                    //----------------------------------------------------------- Viernes ----------------------------------------------------------------------//
                                    EntradaTARDE = Convert.ToDateTime(Viernes[0].ToString()).Date + TARDE;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Viernes[0].ToString()) != null) && (Convert.ToDateTime(Viernes[0].ToString()) < (EntradaTARDE.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), EntradaTARDE);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Viernes[0].ToString()) != null) && (Convert.ToDateTime(Viernes[0].ToString()) >= (EntradaTARDE.AddMinutes(-30))) && (Convert.ToDateTime(Viernes[0].ToString()) < (EntradaTARDE.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), EntradaTARDE);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Viernes[0].ToString()) != null) && (Convert.ToDateTime(Viernes[0].ToString()) > (EntradaTARDE.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), EntradaTARDE);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Viernes[0].ToString(), "TARDE");
                                    }
                                }
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA"))//TURNO SAB
                                {
                                    TimeSpan TDIA = new TimeSpan(07, 00, 0);//Variable con el horario 07:00 hrs
                                    //----------------------------------------------------------- Viernes ----------------------------------------------------------------------//
                                    EntradaTDIA = Convert.ToDateTime(Viernes[0].ToString()).Date + TDIA;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Viernes[0].ToString()) != null) && (Convert.ToDateTime(Viernes[0].ToString()) < (EntradaTDIA.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), EntradaTDIA);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Viernes[0].ToString()) != null) && (Convert.ToDateTime(Viernes[0].ToString())) >= (EntradaTDIA.AddMinutes(-30)) && (Convert.ToDateTime(Viernes[0].ToString()) < (EntradaTDIA.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), EntradaTDIA);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Viernes[0].ToString()) != null) && (Convert.ToDateTime(Viernes[0].ToString()) > (EntradaTDIA.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), EntradaTDIA);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Viernes[0].ToString(), "TDIA");
                                    }
                                }
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ"))//TURNO SAB
                                {
                                    TimeSpan TFIJ = new TimeSpan(08, 00, 0);//Variable con el horario 07:00 hrs
                                    //----------------------------------------------------------- Viernes ----------------------------------------------------------------------//
                                    EntradaTFIJ = Convert.ToDateTime(Viernes[0].ToString()).Date + TFIJ;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Viernes[0].ToString()) != null) && (Convert.ToDateTime(Viernes[0].ToString()) < (EntradaTFIJ.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), EntradaTFIJ);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Viernes[0].ToString()) != null) && (Convert.ToDateTime(Viernes[0].ToString()) >= (EntradaTFIJ.AddMinutes(-30))) && (Convert.ToDateTime(Viernes[0].ToString()) < (EntradaTFIJ.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), EntradaTFIJ);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Viernes[0].ToString()) != null) && (Convert.ToDateTime(Viernes[0].ToString()) > (EntradaTFIJ.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), EntradaTFIJ);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Viernes[0].ToString(), "TFIJ");
                                    }
                                }
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2"))//TURNO SAB
                                {
                                    TimeSpan TNO2 = new TimeSpan(07, 00, 0);//Variable con el horario 07:00 hrs
                                    //----------------------------------------------------------- Viernes ----------------------------------------------------------------------//
                                    EntradaTNO2 = Convert.ToDateTime(Viernes[0].ToString()).Date + TNO2;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Viernes[0].ToString()) != null) && (Convert.ToDateTime(Viernes[0].ToString()) < (EntradaTNO2.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), EntradaTNO2);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Viernes[0].ToString()) != null) && (Convert.ToDateTime(Viernes[0].ToString()) >= (EntradaTNO2.AddMinutes(-30))) && (Convert.ToDateTime(Viernes[0].ToString()) < (EntradaTNO2.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), EntradaTNO2);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Viernes[0].ToString()) != null) && (Convert.ToDateTime(Viernes[0].ToString()) > (EntradaTNO2.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), EntradaTNO2);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Viernes[0].ToString(), "TNO2");
                                    }
                                }
                            } catch (Exception ex) { }
                            break;

                        case "Sábado":
                        case "Saturday":
                            try
                            {
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE"))//TURNO NOCHE
                                {
                                    TimeSpan Noche = new TimeSpan(19, 00, 0);//Variable con el horario 19:00 hrs
                                    //----------------------------------------------------------- Sabado ----------------------------------------------------------------------//
                                    EntradaNoche = Convert.ToDateTime(Sabado[1].ToString()).Date + Noche;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Sabado[1].ToString()) != null) && (Convert.ToDateTime(Sabado[1].ToString()) < (EntradaNoche.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[1].ToString()), EntradaNoche);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Sabado[1].ToString()) != null) && (Convert.ToDateTime(Sabado[1].ToString()) >= (EntradaNoche.AddMinutes(-30))) && (Convert.ToDateTime(Sabado[1].ToString()) < (EntradaNoche.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[1].ToString()), EntradaNoche);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Sabado[1].ToString()) != null) && (Convert.ToDateTime(Sabado[1].ToString()) > (EntradaNoche.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[1].ToString()), EntradaNoche);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Sabado[1].ToString(), "NOCHE");
                                    }
                                }//fin noche
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB"))//SAB
                                {
                                    TimeSpan SAB = new TimeSpan(08, 00, 0);//Variable con el horario 08:00 hrs
                                    //----------------------------------------------------------- Sabado ----------------------------------------------------------------------//
                                    EntradaSAB = Convert.ToDateTime(Sabado[0].ToString()).Date + SAB;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Sabado[0].ToString()) != null) && (Convert.ToDateTime(Sabado[0].ToString()) < (EntradaSAB.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), EntradaSAB);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Sabado[0].ToString()) != null) && (Convert.ToDateTime(Sabado[0].ToString()) >= (EntradaSAB.AddMinutes(-30))) && (Convert.ToDateTime(Sabado[0].ToString()) < (EntradaSAB.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), EntradaSAB);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Sabado[0].ToString()) != null) && (Convert.ToDateTime(Sabado[0].ToString()) > (EntradaSAB.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), EntradaSAB);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Sabado[0].ToString(), "SAB");
                                    }
                                }
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE"))//TARDE
                                {
                                    TimeSpan TARDE = new TimeSpan(15, 30, 0);//Variable con el horario 15:30 hrs
                                    //----------------------------------------------------------- Sabado ----------------------------------------------------------------------//
                                    EntradaTARDE = Convert.ToDateTime(Sabado[0].ToString()).Date + TARDE;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Sabado[0].ToString()) != null) && Convert.ToDateTime(Sabado[0].ToString()) < (EntradaTARDE.AddMinutes(-30)))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), EntradaTARDE);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Sabado[0].ToString()) != null) && (Convert.ToDateTime(Sabado[0].ToString()) >= (EntradaTARDE.AddMinutes(-30))) && (Convert.ToDateTime(Sabado[0].ToString()) < (EntradaTARDE.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), EntradaTARDE);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Sabado[0].ToString()) != null) && (Convert.ToDateTime(Sabado[0].ToString()) > (EntradaTARDE.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), EntradaTARDE);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Sabado[0].ToString(), "TARDE");
                                    }
                                }
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA"))//TDIA
                                {
                                    TimeSpan TDIA = new TimeSpan(07, 00, 0);//Variable con el horario 07:00 hrs
                                    //----------------------------------------------------------- Sabado ----------------------------------------------------------------------//
                                    EntradaTDIA = Convert.ToDateTime(Sabado[0].ToString()).Date + TDIA;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Sabado[0].ToString()) != null) && (Convert.ToDateTime(Sabado[0].ToString()) < (EntradaTDIA.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), EntradaTDIA);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Sabado[0].ToString()) != null) && (Convert.ToDateTime(Sabado[0].ToString()) >= (EntradaTDIA.AddMinutes(-30))) && (Convert.ToDateTime(Sabado[0].ToString()) < (EntradaTDIA.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), EntradaTDIA);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Sabado[0].ToString()) != null) && (Convert.ToDateTime(Sabado[0].ToString()) > (EntradaTDIA.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), EntradaTDIA);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Sabado[0].ToString(), "TDIA");
                                    }
                                }
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ"))//TFIJ
                                {
                                    TimeSpan TFIJ = new TimeSpan(08, 00, 0);//Variable con el horario 07:00 hrs
                                    //----------------------------------------------------------- Sabado ----------------------------------------------------------------------//
                                    EntradaTFIJ = Convert.ToDateTime(Sabado[0].ToString()).Date + TFIJ;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Sabado[0].ToString()) != null) && (Convert.ToDateTime(Sabado[0].ToString()) < (EntradaTFIJ.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), EntradaTFIJ);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Sabado[0].ToString()) != null) && (Convert.ToDateTime(Sabado[0].ToString()) >= (EntradaTFIJ.AddMinutes(-30))) && (Convert.ToDateTime(Sabado[0].ToString()) < (EntradaTFIJ.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), EntradaTFIJ);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Sabado[0].ToString()) != null) && (Convert.ToDateTime(Sabado[0].ToString()) > (EntradaTFIJ.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), EntradaTFIJ);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Sabado[0].ToString(), "TFIJ");
                                    }
                                }
                                else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2"))//TNO2
                                {
                                    TimeSpan TNO2 = new TimeSpan(07, 00, 0);//Variable con el horario 07:00 hrs
                                    //----------------------------------------------------------- Sabado ----------------------------------------------------------------------//
                                    EntradaTNO2 = Convert.ToDateTime(Sabado[0].ToString()).Date + TNO2;//Establece fecha de la checada con la hora de ts
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Sabado[0].ToString()) != null) && (Convert.ToDateTime(Sabado[0].ToString()) < (EntradaTNO2.AddMinutes(-30))))//Entrada de Trabajador a mas de 1 hora del Turno
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), EntradaTNO2);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Sabado[0].ToString()) != null) && (Convert.ToDateTime(Sabado[0].ToString()) >= (EntradaTNO2.AddMinutes(-30))) && (Convert.ToDateTime(Sabado[0].ToString()) < (EntradaTNO2.AddMinutes(6))))
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), EntradaTNO2);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Sabado[0].ToString()) != null) && (Convert.ToDateTime(Sabado[0].ToString()) > (EntradaTNO2.AddMinutes(6))))
                                    {
                                        ///////////////////// ------- FUNCION DESPUES DE 6 MINUTOS -----////////////////////////////////////   
                                        c.LLegadaTarde(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), EntradaTNO2);
                                        RH.ListaRetardos(Empleado[indiceEmpleados].ToString(), Sabado[0].ToString(), "TNO2");
                                    }
                                }
                            } catch (Exception ex) { }
                            break;
                    }//FIN SWITCH dias
                }//FIN FOR Turnos
            }//FIN FOR EMPLEADO
            cn.Close();
            BarraProgreso.Visible = false;
            chekEntrada.Checked = true;
            MessageBox.Show("Registros actualizados");
        }//Fin Funcion_Descontar

        /*****************************************************************************************************************************************************************************************************************************************************************************************************
        ******************************************************************************************************************************************************************************************************************************************************************************************************
        ************************************************************ ========================== C H E C A D O R ============================================= *************************************************************************************************************************************************
        ******************************************************************************************************************************************************************************************************************************************************************************************************
        ******************************************************************************************************************************************************************************************************************************************************************************************************/
        private void butSemanaEmpleado_Click(object sender, EventArgs e)
        {
            this.VerActualizarChecadas();
            gridSemanaTrabajador.Columns["Estatus"].Visible = false;  //Esconde Estatus
           // gridSemanaTrabajador.Columns["HorasTotales"].Visible = false;  //Esconde Estatus
        }//Ver semana de un empleado

        private void VerActualizarChecadas()
        {
            if (txtEmpleado.Text.Equals("#Empleado"))
            {
                MessageBox.Show("No hay empleado a buscar");
            }
            else
            {
                if (gridSemanaTrabajador.Columns.Count != 0)//Limpia El Grid
                {
                    gridSemanaTrabajador.Rows.Clear();
                    gridSemanaTrabajador.Refresh();
                    int i = 0;
                    while (i <= gridSemanaTrabajador.Columns.Count)
                    {
                        gridSemanaTrabajador.Columns.RemoveAt(0);
                        i = +1;
                    }
                }
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////// CONSULTA SEMANA DE LOS TRABAJADORES /////////////////////////////////////////////////////////
                DateTime FechaInicio, FechaFinal;
                String FechaFormateada, FinalFormateada;

                FechaInicio = FechaInicio_Date.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro
                                                     // FechaInicio=FechaInicio.AddDays(-1);//Se agrega 1 dia más
                FechaFormateada = FechaInicio.ToString("yyyy-MM-dd 00:00:00:000");

                FechaFinal = FechaFinal_Date.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro  
                FechaFinal = FechaFinal.AddDays(1);
                FinalFormateada = FechaFinal.ToString("yyyy-MM-dd 00:00:00:000");

                MessageBox.Show("Fecha inicio: " + FechaInicio + " \n A \n fecha final: " + FechaFinal);
                try
                {
                    cn = new SqlConnection(var.GetConexionDB());
                    cn.Open();
                    //importar la libreria de widws form para abrir una ventana 
                    SqlCommand cmd = new SqlCommand("SemanaEmpleado", cn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@FechaInicio", FechaFormateada);
                    cmd.Parameters.AddWithValue("@FechaFinal", FinalFormateada);
                    cmd.Parameters.AddWithValue("@Empleado", txtEmpleado.Text.ToString());
                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.Load(cmd.ExecuteReader());

                    FC.ObetenrDias(dt);//Obtiene los checadas de cada dia
                    var resultado = FC.ObetenrDias(dt);//Contiene la coleccion de ArrayList de la semana
                    ArrayList Lunes = resultado.Item1;
                    ArrayList Martes = resultado.Item2;
                    ArrayList Miercoles = resultado.Item3;
                    ArrayList Jueves = resultado.Item4;
                    ArrayList Viernes = resultado.Item5;
                    ArrayList Sabado = resultado.Item6;

                    this.Columnas(gridSemanaTrabajador);

                    gridSemanaTrabajador.Rows.Add();
                    for (int i = 0; i < 6; i++)
                    {
                        gridSemanaTrabajador.Rows[i].Cells["Trabajador"].Value = dt.Rows[i]["Trabajador"].ToString();
                        gridSemanaTrabajador.Rows[i].Cells["Turno"].Value = dt.Rows[i]["Turno"].ToString();
                        gridSemanaTrabajador.Rows[i].Cells["Estatus"].Value = dt.Rows[i]["Estatus"].ToString();
                        gridSemanaTrabajador.Rows.Add();
                    }
                    for (int elementos = 0; elementos < Lunes.Count; elementos++)
                    {
                        if (Lunes[elementos].ToString() != null)
                        {
                            gridSemanaTrabajador.Rows[elementos].Cells["Lunes"].Value = Convert.ToDateTime(Lunes[elementos]).ToString("yyyy-MM-dd HH:mm:ss");//Se crea fila de [Empleado y tipo checada]
                        }
                    }
                    for (int elementos = 0; elementos < Martes.Count; elementos++)
                    {
                        if (Martes[elementos].ToString() != null)
                        {
                            gridSemanaTrabajador.Rows[elementos].Cells["Martes"].Value = Convert.ToDateTime(Martes[elementos]).ToString("yyyy-MM-dd HH:mm:ss");//Se crea fila de [Empleado y tipo checada]
                        }
                    }
                    for (int elementos = 0; elementos < Miercoles.Count; elementos++)
                    {
                        if (Miercoles[elementos].ToString() != null)
                        {
                            gridSemanaTrabajador.Rows[elementos].Cells["Miércoles"].Value = Convert.ToDateTime(Miercoles[elementos]).ToString("yyyy-MM-dd HH:mm:ss");//Se crea fila de [Empleado y tipo checada]
                        }
                    }
                    for (int elementos = 0; elementos < Jueves.Count; elementos++)
                    {
                        if (Jueves[elementos].ToString() != null)
                        {
                            gridSemanaTrabajador.Rows[elementos].Cells["Jueves"].Value = Convert.ToDateTime(Jueves[elementos]).ToString("yyyy-MM-dd HH:mm:ss");//Se crea fila de [Empleado y tipo checada]
                        }
                    }
                    for (int elementos = 0; elementos < Viernes.Count; elementos++)
                    {
                        if (Viernes[elementos].ToString() != null)
                        {
                            gridSemanaTrabajador.Rows[elementos].Cells["Viernes"].Value = Convert.ToDateTime(Viernes[elementos]).ToString("yyyy-MM-dd HH:mm:ss");//Se crea fila de [Empleado y tipo checada]
                        }
                    }
                    for (int elementos = 0; elementos < Sabado.Count; elementos++)
                    {
                        if (Sabado[elementos].ToString() != null)
                        {
                            gridSemanaTrabajador.Rows[elementos].Cells["Sábado"].Value = Convert.ToDateTime(Sabado[elementos]).ToString("yyyy-MM-dd HH:mm:ss");//Se crea fila de [Empleado y tipo checada]
                        }
                    }
                    gridIncidencias.DataSource = gridSemanaTrabajador.DataSource;
                }
                catch (Exception ex) { MessageBox.Show("Error" + ex); }
            }
        }

        private void butAplicar_Click(object sender, EventArgs e)//Crea una nueva checada faltante//
        {
            String TipoChecada = listTipoChecada.Text;
            //******************************************* FORMATO DE CHECADA **********************************************************************************/
            String Horas = txtHoraNueva.Text;
            int Hora = Convert.ToInt32(Horas.Substring(0, 2));
            int Minutos = Convert.ToInt32(Horas.Substring(3, 2));
            int Segundos = Convert.ToInt32(0/*Horas.Substring(6,2)*/);
            DateTime NuevaFecha;
            DateTime NuevaChecada;
            String FechaFormateada;
            NuevaFecha = dateChecadaNueva.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro
            FechaFormateada = NuevaFecha.ToString("yyyy-MM-dd 00:00:0");
            TimeSpan TS_Horas = new TimeSpan(Hora, Minutos, Segundos);//Variable con el horario 19:00 hrs
            NuevaChecada = Convert.ToDateTime(Convert.ToDateTime(FechaFormateada) + TS_Horas);//Establece fecha de la checada con la hora de ts
            //**************************************************************************************************************************************************/

            DialogResult result = MessageBox.Show("Insertar Checada: " + NuevaChecada.ToShortDateString(), "Salir", MessageBoxButtons.YesNoCancel);
            switch (result)
            {
                case DialogResult.Yes:
                    FC.InsertarChecada(NuevaChecada, txtEmpleado.Text, listTipoChecada.Text.ToString(), listTurno.Text.ToString(), "Error", "Checada");
                    this.VerActualizarChecadas();
                    break;
                case DialogResult.No:
                    break;
                case DialogResult.Cancel:
                    break;
            }
        }

        private void txtHoras_Enter(object sender, EventArgs e)
        {
            txtHoraNueva.Text = "";
        }

        private void txtEmpleado_Enter(object sender, EventArgs e)
        {
            txtEmpleado.Text = "";
        }

        private void txtEmpleado_Leave(object sender, EventArgs e)
        {
            if (txtEmpleado.Text.Equals(""))
            {
                txtEmpleado.Text = "#Empleado";
            }
        }

        private void butEliminarChecada_Click(object sender, EventArgs e)
        {
            String Turno = gridSemanaTrabajador.Rows[gridSemanaTrabajador.CurrentRow.Index].Cells["Turno"].Value.ToString();
            FC.EliminarChecada(Convert.ToDateTime(gridSemanaTrabajador.CurrentCell.Value), txtEmpleado.Text.ToString(), Turno);
            MessageBox.Show("Empleado: " + txtEmpleado.Text.ToString() + " Turno: " + Turno + " Checada: " + gridSemanaTrabajador.CurrentCell.Value.ToString());
            this.VerActualizarChecadas();
        }

        private void txtHoras_Leave(object sender, EventArgs e)
        {
            if (txtHoraNueva.Text.Equals(""))
            {
                txtHoraNueva.Text = "23:59:00";
            }
        }

        private void monoFlat_Button3_Click(object sender, EventArgs e)
        {
            ListaErrores("gridErrores");
        }

        private void ListaErrores(String grid)
        {
            ArrayList Empleado = new ArrayList();
            /////////////////////////////////////GET COLECCIÓN && SELECCIÓN DE EMPLEADOS/////////////////////////////////////////////////////////////
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            query = "SELECT DISTINCT Trabajador FROM Checadas WHERE Estatus = 'Error' Order by Trabajador ";

            SqlCommand cmd2 = new SqlCommand(query, cn);
            DataTable t1 = new DataTable();
            using (SqlDataAdapter a = new SqlDataAdapter(cmd2))
            {
                a.Fill(t1);
            }
            cn.Close();

            if (grid.Equals("gridErrores"))
            {
                gridErrores.Rows.Clear();
                gridErrores.DataSource = t1;
            }
        }

        private void butFaltas_Click(object sender, EventArgs e)
        {
            ArrayList Lunes = new ArrayList();
            ArrayList Martes = new ArrayList();
            ArrayList Miercoles = new ArrayList();
            ArrayList Jueves = new ArrayList();
            ArrayList Viernes = new ArrayList();
            ArrayList Sabado = new ArrayList();
            ArrayList Semana = new ArrayList { "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado" };
            ArrayList Empleado = new ArrayList();

            /////////////////////////////////////GET COLECCIÓN && SELECCIÓN DE EMPLEADOS POR TURNOS /////////////////////////////////////////////////////////////
            Empleado.Clear();
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            string query = "SELECT DISTINCT Trabajador FROM Checadas Order by Trabajador ";

            SqlCommand cmd = new SqlCommand(query, cn);
            DataTable t1 = new DataTable();
            using (SqlDataAdapter a = new SqlDataAdapter(cmd))
            {
                a.Fill(t1);
            }

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

                String TurnoTrabajador = dt.Rows[0]["Turno"].ToString();
                /////////////////////////////////////////////// OBTENER CHECADAS DEL DIA //////////////////////////////////////////////////////
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

                DateTime inicio = dateInicio.Value;
                for (int sem = 0; sem < Semana.Count; sem++)
                {
                    switch (Semana[sem])
                    {
                        case "Lunes":
                            int registrosLunes = 0;
                            for (int Lun = 0; Lun <= Lunes.Count; Lun++)
                            {
                                if (Lunes.Count == 0)
                                {
                                    if (TurnoTrabajador.Equals("Noche"))
                                    {
                                        while (registrosLunes < 1)
                                        {
                                            Lunes.Add(inicio);
                                            //MessageBox.Show("Falta en lunes: " + Empleado[indiceEmpleados].ToString());
                                            registrosLunes++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TARDE"))
                                    {
                                        while (registrosLunes < 2)
                                        {
                                            Lunes.Add(inicio);
                                            //MessageBox.Show("Falta en lunes: " + Empleado[indiceEmpleados].ToString());
                                            registrosLunes++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TDIA"))
                                    {
                                        while (registrosLunes < 4)
                                        {
                                            Lunes.Add(inicio);
                                            //MessageBox.Show("Falta en lunes: " + Empleado[indiceEmpleados].ToString());
                                            registrosLunes++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TFIJ"))
                                    {
                                        while (registrosLunes < 4)
                                        {
                                            Lunes.Add(inicio);
                                            //MessageBox.Show("Falta en lunes: " + Empleado[indiceEmpleados].ToString());
                                            registrosLunes++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TNO2"))
                                    {
                                        while (registrosLunes < 4)
                                        {
                                            Lunes.Add(inicio);
                                            //MessageBox.Show("Falta en lunes: " + Empleado[indiceEmpleados].ToString());
                                            registrosLunes++;
                                        }
                                    }
                                }//IF Lunes
                            }//For Lunes
                            break;
                        case "Martes":
                            int registrosMartes = 0;
                            for (int Mar = 0; Mar <= Martes.Count; Mar++)
                            {
                                if (Martes.Count == 0)
                                {
                                    if (TurnoTrabajador.Equals("Noche"))
                                    {
                                        while (registrosMartes < 2)
                                        {
                                            Martes.Add(inicio);
                                            //MessageBox.Show("Falta en Martes: " + Empleado[indiceEmpleados].ToString());
                                            registrosMartes++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TARDE"))
                                    {
                                        while (registrosMartes < 2)
                                        {
                                            Martes.Add(inicio);
                                            // MessageBox.Show("Falta en Martes: " + Empleado[indiceEmpleados].ToString());
                                            registrosMartes++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TDIA"))
                                    {
                                        while (registrosMartes < 4)
                                        {
                                            Martes.Add(inicio);
                                            //MessageBox.Show("Falta en Martes: " + Empleado[indiceEmpleados].ToString());
                                            registrosMartes++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TFIJ"))
                                    {
                                        while (registrosMartes < 4)
                                        {
                                            Martes.Add(inicio);
                                            // MessageBox.Show("Falta en Martes: " + Empleado[indiceEmpleados].ToString());
                                            registrosMartes++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TNO2"))
                                    {
                                        while (registrosMartes < 4)
                                        {
                                            Martes.Add(inicio);
                                            // MessageBox.Show("Falta en Martes: " + Empleado[indiceEmpleados].ToString());
                                            registrosMartes++;
                                        }
                                    }
                                }//IF Martes
                            }//For Martes
                            break;
                        case "Miercoles":
                            int registrosMiercoles = 0;
                            for (int Mie = 0; Mie <= Miercoles.Count; Mie++)
                            {
                                if (Miercoles.Count == 0)
                                {
                                    if (TurnoTrabajador.Equals("Noche"))
                                    {
                                        while (registrosMiercoles < 2)
                                        {
                                            Miercoles.Add(inicio);
                                            //MessageBox.Show("Falta en Miercoles: " + Empleado[indiceEmpleados].ToString());
                                            registrosMiercoles++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TARDE"))
                                    {
                                        while (registrosMiercoles < 2)
                                        {
                                            Miercoles.Add(inicio);
                                            //MessageBox.Show("Falta en Miercoles: " + Empleado[indiceEmpleados].ToString());
                                            registrosMiercoles++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TDIA"))
                                    {
                                        while (registrosMiercoles < 4)
                                        {
                                            Miercoles.Add(inicio);
                                            // MessageBox.Show("Falta en Miercoles: " + Empleado[indiceEmpleados].ToString());
                                            registrosMiercoles++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TFIJ"))
                                    {
                                        while (registrosMiercoles < 4)
                                        {
                                            Miercoles.Add(inicio);
                                            //MessageBox.Show("Falta en Miercoles: " + Empleado[indiceEmpleados].ToString());
                                            registrosMiercoles++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TNO2"))
                                    {
                                        while (registrosMiercoles < 4)
                                        {
                                            Miercoles.Add(inicio);
                                            //MessageBox.Show("Falta en Miercoles: " + Empleado[indiceEmpleados].ToString());
                                            registrosMiercoles++;
                                        }
                                    }
                                }//IF Martes
                            }//For Martes
                            break;
                        case "Jueves":
                            int registrosJueves = 0;
                            for (int Jue = 0; Jue <= Jueves.Count; Jue++)
                            {
                                if (Jueves.Count == 0)
                                {
                                    if (TurnoTrabajador.Equals("Noche"))
                                    {
                                        while (registrosJueves < 2)
                                        {
                                            Jueves.Add(inicio);
                                            //MessageBox.Show("Falta en Jueves: " + Empleado[indiceEmpleados].ToString());
                                            registrosJueves++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TARDE"))
                                    {
                                        while (registrosJueves < 2)
                                        {
                                            Jueves.Add(inicio);
                                            // MessageBox.Show("Falta en Jueves: " + Empleado[indiceEmpleados].ToString());
                                            registrosJueves++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TDIA"))
                                    {
                                        while (registrosJueves < 4)
                                        {
                                            Jueves.Add(inicio);
                                            //MessageBox.Show("Falta en Jueves: " + Empleado[indiceEmpleados].ToString());
                                            registrosJueves++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TFIJ"))
                                    {
                                        while (registrosJueves < 4)
                                        {
                                            Jueves.Add(inicio);
                                            // MessageBox.Show("Falta en Jueves: " + Empleado[indiceEmpleados].ToString());
                                            registrosJueves++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TNO2"))
                                    {
                                        while (registrosJueves < 4)
                                        {
                                            Jueves.Add(inicio);
                                            //MessageBox.Show("Falta en Jueves: " + Empleado[indiceEmpleados].ToString());
                                            registrosJueves++;
                                        }
                                    }
                                }//IF Martes
                            }//For Martes
                            break;
                        case "Viernes":
                            int registrosViernes = 0;
                            for (int Vie = 0; Vie <= Viernes.Count; Vie++)
                            {
                                if (Viernes.Count == 0)
                                {
                                    if (TurnoTrabajador.Equals("Noche"))
                                    {
                                        while (registrosViernes < 2)
                                        {
                                            Viernes.Add(inicio);
                                            //MessageBox.Show("Falta en Viernes: " + Empleado[indiceEmpleados].ToString());
                                            registrosViernes++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TARDE"))
                                    {
                                        while (registrosViernes < 2)
                                        {
                                            Viernes.Add(inicio);
                                            //MessageBox.Show("Falta en Viernes: " + Empleado[indiceEmpleados].ToString());
                                            registrosViernes++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TDIA"))
                                    {
                                        while (registrosViernes < 4)
                                        {
                                            Viernes.Add(inicio);
                                            //MessageBox.Show("Falta en Viernes: " + Empleado[indiceEmpleados].ToString());
                                            registrosViernes++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TFIJ"))
                                    {
                                        while (registrosViernes < 4)
                                        {
                                            Viernes.Add(inicio);
                                            // MessageBox.Show("Falta en Viernes: " + Empleado[indiceEmpleados].ToString());
                                            registrosViernes++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TNO2"))
                                    {
                                        while (registrosViernes < 4)
                                        {
                                            Viernes.Add(inicio);
                                            //MessageBox.Show("Falta en Viernes: " + Empleado[indiceEmpleados].ToString());
                                            registrosViernes++;
                                        }
                                    }
                                }//IF Martes
                            }//For Martes
                            break;
                        case "Sabado":
                            int registrosSabado = 0;
                            for (int Sab = 0; Sab <= Sabado.Count; Sab++)
                            {
                                if (Sabado.Count == 0)
                                {
                                    if (TurnoTrabajador.Equals("Noche"))
                                    {
                                        while (registrosSabado < 2)
                                        {
                                            Viernes.Add(inicio);
                                            //MessageBox.Show("Falta en Viernes: " + Empleado[indiceEmpleados].ToString());
                                            registrosSabado++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TARDE"))
                                    {
                                        while (registrosSabado < 2)
                                        {
                                            Viernes.Add(inicio);
                                            // MessageBox.Show("Falta en Viernes: " + Empleado[indiceEmpleados].ToString());
                                            registrosSabado++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TDIA"))
                                    {
                                        while (registrosSabado < 4)
                                        {
                                            Viernes.Add(inicio);
                                            //MessageBox.Show("Falta en Viernes: " + Empleado[indiceEmpleados].ToString());
                                            registrosSabado++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TFIJ"))
                                    {
                                        while (registrosSabado < 4)
                                        {
                                            Viernes.Add(inicio);
                                            //MessageBox.Show("Falta en Viernes: " + Empleado[indiceEmpleados].ToString());
                                            registrosSabado++;
                                        }
                                    }
                                    else if (TurnoTrabajador.Equals("TNO2"))
                                    {
                                        while (registrosSabado < 4)
                                        {
                                            Viernes.Add(inicio);
                                            //MessageBox.Show("Falta en Viernes: " + Empleado[indiceEmpleados].ToString());
                                            registrosSabado++;
                                        }
                                    }
                                }//IF Martes
                            }//For Martes
                            break;
                    }//Switch de dia de la semana
                }//For List dias de la semana
            }
        }

        private void butInsertarFalta_Click(object sender, EventArgs e)
        {
            String TipoChecada = listTipoChecada.Text;
            //******************************************* FORMATO DE CHECADA **********************************************************************************/
            String Horas = txtHoraNueva.Text;
            int Hora = Convert.ToInt32(Horas.Substring(0, 2));
            int Minutos = Convert.ToInt32(Horas.Substring(3, 2));
            int Segundos = Convert.ToInt32(0/*Horas.Substring(6,2)*/);
            DateTime NuevaFecha;
            DateTime NuevaChecada;
            String FechaFormateada;
            NuevaFecha = dateChecadaNueva.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro
            FechaFormateada = NuevaFecha.ToString("yyyy-MM-dd 00:00:0");

            TimeSpan TS_Horas = new TimeSpan(Hora, Minutos, Segundos);//Variable con el horario 19:00 hrs
            NuevaChecada = Convert.ToDateTime(Convert.ToDateTime(FechaFormateada) + TS_Horas);//Establece fecha de la checada con la hora de ts
            //**************************************************************************************************************************************************/
            if (listTurno.Text.Equals("NOCHE"))
            {
                int limite = 0;
                while (limite < NOCHE)
                {

                    FC.InsertarChecada(NuevaChecada, txtEmpleado.Text, listTipoChecada.Text.ToString(), listTurno.Text.ToString(), "Error", "Falta");
                    limite++;
                }
            }
            if (listTurno.Text.Equals("TNO2"))
            {
                int limite = 0;
                while (limite < TNO2)
                {
                    FC.InsertarChecada(NuevaChecada, txtEmpleado.Text, listTipoChecada.Text.ToString(), listTurno.Text.ToString(), "Error", "Falta");
                    limite++;
                }
            }
            if (listTurno.Text.Equals("TARDE"))
            {
                int limite = 0;
                while (limite < TARDE)
                {
                    FC.InsertarChecada(NuevaChecada, txtEmpleado.Text, listTipoChecada.Text.ToString(), listTurno.Text.ToString(), "Error", "Falta");
                    limite++;
                }
            }
            if (listTurno.Text.Equals("TDIA"))
            {
                int limite = 0;
                while (limite < TDIA)
                {
                    FC.InsertarChecada(NuevaChecada, txtEmpleado.Text, listTipoChecada.Text.ToString(), listTurno.Text.ToString(), "Error", "Falta");
                    limite++;
                }
            }
            if (listTurno.Text.Equals("TFIJ"))
            {
                int limite = 0;
                while (limite < TFIJ)
                {
                    FC.InsertarChecada(NuevaChecada, txtEmpleado.Text, listTipoChecada.Text.ToString(), listTurno.Text.ToString(), "Error", "Falta");
                    limite++;
                }
            }
        }

        private void pruebas_Click(object sender, EventArgs e)
        {
            /////////////////////////////////////////////////////////////////////////////////////////////////////////
            //////================================ OBTENER LISTA DE EMPLEADOS ===============================////////
            /////////////////////////////////////////////////////////////////////////////////////////////////////////
            if (GridSemana.Rows.Count > 1)
            {
                for (int i = GridSemana.Rows.Count - 1; i >= 0; i--)
                {
                    GridSemana.Rows.RemoveAt(i);
                }
            }
            GridSemana.Refresh();
            ArrayList Lunes = new ArrayList();
            ArrayList Martes = new ArrayList();
            ArrayList Miercoles = new ArrayList();
            ArrayList Jueves = new ArrayList();
            ArrayList Viernes = new ArrayList();
            ArrayList Sabado = new ArrayList();
            ArrayList Empleado = new ArrayList();
            ArrayList Entrada = new ArrayList();
            ArrayList Salida = new ArrayList();
            ArrayList Turno = new ArrayList() { "TDIA", "TFIJ", "TNO2", "TARDE"};
            ArrayList Semana = new ArrayList() { "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado" };

            BarraProgreso.Maximum = Turno.Count;
            BarraProgreso.BackColor = Color.White;
            BarraProgreso.ProgressColor1 = Color.CornflowerBlue;
            BarraProgreso.ProgressColor2 = Color.DarkRed;
            BarraProgreso.Visible = true;
            BarraProgreso.ForeColor = Color.Cyan;
            try
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                cmd = new SqlCommand("TRUNCATE TABLE SemanaNomina", cn);
                cmd.ExecuteNonQuery();//para que relice la sentencia sql
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se logro limpiar tabla SemanaNomina" + ex.Message.ToString());
            }

            for (int turn = 0; turn < 4; turn++)
            {
                BarraProgreso.Value = turn;
                BarraProgreso.Update();//Contador de progreso
                Empleado.Clear();
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                query = "SELECT DISTINCT Trabajador FROM Checadas WHERE Turno = '" + Turno[turn].ToString() + "' Order by Trabajador";
                //MessageBox.Show(Turno[turn].ToString());
                SqlCommand cmd = new SqlCommand(query, cn);
                DataTable dt = new DataTable();
                using (SqlDataAdapter a = new SqlDataAdapter(cmd))
                {
                    a.Fill(dt);
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    Empleado.Add(dt.Rows[i]["Trabajador"]);
                    //   MessageBox.Show(Empleado[i].ToString());
                }
                cn.Close();
               
                /////////////////////////////////////////////////////////////////////////////////////////////////////////
                //////======================== OBTENER REGISTROS SEMANAL DE 1 EMPLEADO ==========================////////
                /////////////////////////////////////////////////////////////////////////////////////////////////////////
                DateTime FechaInicio, FechaFinal;
                String FechaFormateada, FinalFormateada;

                FechaInicio = dateInicio.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro
                                               // FechaInicio=FechaInicio.AddDays(-1);//Se agrega 1 dia más
                FechaFormateada = FechaInicio.ToString("yyyy-MM-dd 00:00:00:000");

                FechaFinal = dateFinal.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro  
                FechaFinal = FechaFinal.AddDays(1);
                FinalFormateada = FechaFinal.ToString("yyyy-MM-dd 00:00:00:000");

                for (int t = 0; t < Empleado.Count; t++)//Recorre lista de empleados
                {
                    cn = new SqlConnection(var.GetConexionDB());
                    cn.Open();
                    //importar la libreria de widws form para abrir una ventana 
                    cmd = new SqlCommand("SemanaTrabajador", cn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@Empleado", Empleado[t].ToString());
                    cmd.Parameters.AddWithValue("@FechaInicio", FechaFormateada);
                    cmd.Parameters.AddWithValue("@FechaFinal", FinalFormateada);
                    // cmd.Parameters.AddWithValue("@Turno", Turno[turn].ToString());
                    dt.Clear();//Limpia antes de llenar
                    dt.Load(cmd.ExecuteReader());
                    cn.Close();

                    /////////////////////////////////////////////////////////////////////////////////////////////////////////
                    //////============================== OBTENER REGISTROS DE CADA DIA ==============================////////
                    /////////////////////////////////////////////////////////////////////////////////////////////////////////

                    Lunes.Clear();
                    Martes.Clear();
                    Miercoles.Clear();
                    Jueves.Clear();
                    Viernes.Clear();
                    Viernes.Clear();
                    Sabado.Clear();

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
                                if (dt.Rows.Count > 0)
                                {
                                    if (dt.Rows[reg]["Turno"].ToString().Equals("TDIA") || dt.Rows[reg]["Turno"].ToString().Equals("TFIJ") || dt.Rows[reg]["Turno"].ToString().Equals("TNO2") || dt.Rows[reg]["Turno"].ToString().Equals("TARDE"))
                                    {
                                        Lunes.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                    }
                                }
                                break;
                            case "Martes":
                            case "Tuesday":
                                if (dt.Rows.Count > 0)
                                {
                                    if (dt.Rows[reg]["Turno"].ToString().Equals("TDIA") || dt.Rows[reg]["Turno"].ToString().Equals("TFIJ") || dt.Rows[reg]["Turno"].ToString().Equals("TNO2") || dt.Rows[reg]["Turno"].ToString().Equals("TARDE"))
                                    {
                                        Martes.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                    }
                                }
                                break;
                            case "Miércoles":
                            case "Wednesday":
                                if (dt.Rows.Count > 0)
                                {
                                    if (dt.Rows[reg]["Turno"].ToString().Equals("TDIA") || dt.Rows[reg]["Turno"].ToString().Equals("TFIJ") || dt.Rows[reg]["Turno"].ToString().Equals("TNO2") || dt.Rows[reg]["Turno"].ToString().Equals("TARDE"))
                                    {
                                        Miercoles.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                    }
                                }
                                break;
                            case "Jueves":
                            case "Thursday":
                                if (dt.Rows.Count > 0)
                                {
                                    if (dt.Rows[reg]["Turno"].ToString().Equals("TDIA") || dt.Rows[reg]["Turno"].ToString().Equals("TFIJ") || dt.Rows[reg]["Turno"].ToString().Equals("TNO2") || dt.Rows[reg]["Turno"].ToString().Equals("TARDE"))
                                    {
                                        Jueves.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                    }
                                }
                                break;
                            case "Viernes":
                            case "Friday":
                                if (dt.Rows.Count > 0)
                                {
                                    if (dt.Rows[reg]["Turno"].ToString().Equals("TDIA") || dt.Rows[reg]["Turno"].ToString().Equals("TFIJ") || dt.Rows[reg]["Turno"].ToString().Equals("TNO2") || dt.Rows[reg]["Turno"].ToString().Equals("TARDE"))
                                    {
                                        Viernes.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                    }
                                }
                                break;
                            case "Sábado":
                            case "Saturday":
                                if (dt.Rows.Count > 0)
                                {
                                    if (dt.Rows[reg]["Turno"].ToString().Equals("TDIA") || dt.Rows[reg]["Turno"].ToString().Equals("TFIJ") || dt.Rows[reg]["Turno"].ToString().Equals("TNO2") || dt.Rows[reg]["Turno"].ToString().Equals("TARDE"))
                                    {
                                        Sabado.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                    }
                                }
                                break;
                        }//FIN SWITCH dias
                    }//FOR DIAS DE LA SEMANA
                    DateTime SalidaComer;// new DateTime();
                    DateTime RegresoComer;// = new DateTime();
                    DateTime EntradaPotencial; ;// = SalidaComer.AddMinutes(30);

                    for (int i = 0; i < Semana.Count; i++)
                    {
                        switch (Semana[i].ToString())
                        {
                            case "Lunes":
                                if (Lunes.Count > 0)
                                {
                                    try
                                    {
                                        SalidaComer = Convert.ToDateTime(Lunes[1].ToString());
                                        RegresoComer = Convert.ToDateTime(Lunes[2].ToString());
                                        EntradaPotencial = Convert.ToDateTime(Lunes[1].ToString()).AddMinutes(30);

                                        if (RegresoComer < EntradaPotencial.AddMinutes(6))
                                        {
                                            c.ActualizarChecadaComida(Empleado[t].ToString(), SalidaComer, RegresoComer, EntradaPotencial);//Actualiza la checada al turno asignado
                                        }
                                        else
                                        {
                                            c.ComidaTarde(Empleado[t].ToString(), SalidaComer, RegresoComer, EntradaPotencial, Turno[turn].ToString());
                                        }
                                    }
                                    catch (Exception ex) { }
                                }
                                break;
                            case "Martes":
                                if (Martes.Count > 0)
                                {
                                    try
                                    {
                                        SalidaComer = Convert.ToDateTime(Martes[1].ToString());
                                        RegresoComer = Convert.ToDateTime(Martes[2].ToString());
                                        EntradaPotencial = Convert.ToDateTime(Martes[1].ToString()).AddMinutes(30);

                                        if (RegresoComer < EntradaPotencial.AddMinutes(6))
                                        {
                                            c.ActualizarChecadaComida(Empleado[t].ToString(), SalidaComer, RegresoComer, EntradaPotencial);//Actualiza la checada al turno asignado
                                        }
                                        else
                                        {
                                            c.ComidaTarde(Empleado[t].ToString(), SalidaComer, RegresoComer, EntradaPotencial, Turno[turn].ToString());
                                        }
                                    }
                                    catch (Exception ex) { }
                                }
                                break;
                            case "Miercoles":
                                if (Miercoles.Count > 0)
                                {
                                    try
                                    {
                                        SalidaComer = Convert.ToDateTime(Miercoles[1].ToString());
                                        RegresoComer = Convert.ToDateTime(Miercoles[2].ToString());
                                        EntradaPotencial = Convert.ToDateTime(Miercoles[1].ToString()).AddMinutes(30);

                                        if (RegresoComer < EntradaPotencial.AddMinutes(6))
                                        {
                                            c.ActualizarChecadaComida(Empleado[t].ToString(), SalidaComer, RegresoComer, EntradaPotencial);//Actualiza la checada al turno asignado
                                        }
                                        else
                                        {
                                            c.ComidaTarde(Empleado[t].ToString(), SalidaComer, RegresoComer, EntradaPotencial, Turno[turn].ToString());
                                        }
                                    }
                                    catch (Exception ex) { }
                                }
                                break;
                            case "Jueves":
                                if (Jueves.Count > 0)
                                {
                                    try
                                    {
                                        SalidaComer = Convert.ToDateTime(Jueves[1].ToString());
                                        RegresoComer = Convert.ToDateTime(Jueves[2].ToString());
                                        EntradaPotencial = Convert.ToDateTime(Jueves[1].ToString()).AddMinutes(30);

                                        if (RegresoComer < EntradaPotencial.AddMinutes(6))
                                        {
                                            c.ActualizarChecadaComida(Empleado[t].ToString(), SalidaComer, RegresoComer, EntradaPotencial);//Actualiza la checada al turno asignado
                                        }
                                        else
                                        {
                                            c.ComidaTarde(Empleado[t].ToString(), SalidaComer, RegresoComer, EntradaPotencial, Turno[turn].ToString());
                                        }
                                    }
                                    catch (Exception ex) { }
                                }
                                break;
                            case "Viernes":
                                if (Viernes.Count > 0)
                                {
                                    try
                                    {
                                        SalidaComer = Convert.ToDateTime(Viernes[1].ToString());
                                        RegresoComer = Convert.ToDateTime(Viernes[2].ToString());
                                        EntradaPotencial = Convert.ToDateTime(Viernes[1].ToString()).AddMinutes(30);

                                        if (RegresoComer < EntradaPotencial.AddMinutes(6))
                                        {
                                            c.ActualizarChecadaComida(Empleado[t].ToString(), SalidaComer, RegresoComer, EntradaPotencial);//Actualiza la checada al turno asignado
                                        }
                                        else
                                        {
                                            c.ComidaTarde(Empleado[t].ToString(), SalidaComer, RegresoComer, EntradaPotencial, Turno[turn].ToString());
                                        }
                                    }
                                    catch (Exception ex) { }
                                }
                                break;
                            case "Sabado":
                                if (Sabado.Count > 0)
                                {
                                    try
                                    {
                                        SalidaComer = Convert.ToDateTime(Sabado[1].ToString());
                                        RegresoComer = Convert.ToDateTime(Sabado[2].ToString());
                                        EntradaPotencial = Convert.ToDateTime(Sabado[1].ToString()).AddMinutes(30);

                                        if (RegresoComer < EntradaPotencial.AddMinutes(6))
                                        {
                                            c.ActualizarChecadaComida(Empleado[t].ToString(), SalidaComer, RegresoComer, EntradaPotencial);//Actualiza la checada al turno asignado
                                        }
                                        else
                                        {
                                            c.ComidaTarde(Empleado[t].ToString(), SalidaComer, RegresoComer, EntradaPotencial, Turno[turn].ToString());
                                        }
                                    }
                                    catch (Exception ex)
                                    { //MessageBox.Show("Error en empleado " + Empleado + " Por: " + ex); 
                                    }
                                }
                                break;
                        }
                    }
                }//FOR LISTA DE EMPLEADOS
            }
            BarraProgreso.Visible = false;
            chekComida.Checked = true;
            MessageBox.Show("Comidas actualizadas");
        }

        private void butSalidas_Click(object sender, EventArgs e)
        {
           if (GridSemana.Rows.Count > 1)
            {
                for (int i = GridSemana.Rows.Count - 1; i >= 0; i--)
                {
                    GridSemana.Rows.RemoveAt(i);
                }
            }
            GridSemana.Refresh();
            DateTime SalidaNoche = Convert.ToDateTime("07:00:00");
            DateTime SalidaSAB = Convert.ToDateTime("13:00:00");
            // DateTime SalidaTARDE = Convert.ToDateTime("23:00:00");
            DateTime SalidaTDIA = Convert.ToDateTime("15:30:00");
            DateTime SalidaTFIJ = Convert.ToDateTime("17:00:00");
            DateTime SalidaTNO2 = Convert.ToDateTime("19:00:00");

            ArrayList Lunes = new ArrayList();
            ArrayList Martes = new ArrayList();
            ArrayList Miercoles = new ArrayList();
            ArrayList Jueves = new ArrayList();
            ArrayList Viernes = new ArrayList();
            ArrayList Sabado = new ArrayList();
            ArrayList Semana = new ArrayList {"Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado"};
            ArrayList Empleado = new ArrayList();

            /////////////////////////////////////GET COLECCIÓN && SELECCIÓN DE EMPLEADOS POR TURNOS /////////////////////////////////////////////////////////////
            Empleado.Clear();
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            string query = "SELECT DISTINCT Trabajador FROM Checadas Order by Trabajador ";

            SqlCommand cmd = new SqlCommand(query, cn);
            DataTable t1 = new DataTable();
            using (SqlDataAdapter a = new SqlDataAdapter(cmd))
            {
                a.Fill(t1);
            }

            for (int x = 0; x < t1.Rows.Count; x++)//FOR[1] : Agrega los Trabajadores a una lista
            {
                Empleado.Add(t1.Rows[x]["Trabajador"].ToString());
            }

            BarraProgreso.Maximum = Empleado.Count;
            BarraProgreso.BackColor = Color.White;
            BarraProgreso.ProgressColor1 = Color.CornflowerBlue;
            BarraProgreso.ProgressColor2 = Color.DarkRed;
            BarraProgreso.Visible = true;
            BarraProgreso.ForeColor = Color.Cyan;
            ///////////////////////////////////////////// OBTENER DATOS DE CADA EMPLEADO EN LISTA //////////////////////////////////////////////////////
            for (int indiceEmpleados = 0; indiceEmpleados < Empleado.Count; indiceEmpleados++)
            {
                BarraProgreso.Value = indiceEmpleados;
                BarraProgreso.Update();//Contador de progreso
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

                String TurnoTrabajador = dt.Rows[0]["Turno"].ToString();
                /////////////////////////////////////////////// OBTENER CHECADAS DEL DIA //////////////////////////////////////////////////////
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

                /////////////////////////////////////////////// REVISAR SALIDA DIARIA //////////////////////////////////////////////////////

                for (int indiceTurno = 0; indiceTurno < dt.Rows.Count; indiceTurno++)//Recorrer cada registro para obtener el turno
                {
                    ////////////////////////////////// Codigo para el dia de la semana /////////////////////////////////////////////////////
                    DateTime dateValue = new DateTime();
                    CultureInfo cul = CultureInfo.CurrentCulture;
                    // Usa la fecha formateada y calcula el número de la semana
                    int NumeroSemana = cul.Calendar.GetWeekOfYear(dateValue, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
                    dateValue = Convert.ToDateTime(dt.Rows[indiceTurno]["Checada"].ToString());//Consulto la checada del indice
                    String dia = dateValue.DayOfWeek.ToString();//Convierto la fecha en dia de la semana
                    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                    String TurnoRegistro = dt.Rows[indiceTurno]["Turno"].ToString();
                    switch (dia)
                    {
                        case "Lunes":
                        case "Monday":
                            try
                            {
                                DateTime m = Convert.ToDateTime(Lunes[0].ToString());
                                ///////////// - NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE /////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE"))//TURNO NOCHE
                                {
                                    TimeSpan Noche = new TimeSpan(07, 00, 00);//Variable con el horario 19:00 hrs
                                    SalidaNoche = Convert.ToDateTime(Lunes[0].ToString()).Date + Noche;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Lunes ----------------------------------------------------------------------//
                                    if (Lunes.Count == 3)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Lunes[1].ToString()) != null) && (Convert.ToDateTime(Lunes[1].ToString()) >= SalidaNoche && Convert.ToDateTime(Lunes[1].ToString()) < SalidaNoche.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[1].ToString()), SalidaNoche);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Lunes[1].ToString()) != null) && (Convert.ToDateTime(Lunes[1].ToString()) < SalidaNoche))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[1].ToString()), SalidaNoche);
                                        }
                                    }
                                }//Fin noche
                                //////////// - TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE///////////////////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE"))//TURNO NOCHE
                                {
                                    TimeSpan Tarde = new TimeSpan(23, 00, 00);//Variable con el horario 19:00 hrs
                                    DateTime SalidaTarde = Convert.ToDateTime(Lunes[0].ToString()).Date + Tarde;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Lunes ----------------------------------------------------------------------//
                                    if (Lunes.Count == 2)//En caso de 2 checadas
                                    {
                                        if (Convert.ToDateTime(Lunes[1].ToString()) >= SalidaTarde && Convert.ToDateTime(Lunes[1].ToString()) < SalidaTarde.AddMinutes(30))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[1].ToString()), SalidaTarde);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Lunes[1].ToString()) != null) && (Convert.ToDateTime(Lunes[1].ToString()) < SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), SalidaTarde);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Lunes[1].ToString()) != null) && (Convert.ToDateTime(Lunes[1].ToString()) > SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[0].ToString()), SalidaTarde);
                                        }
                                    }//IF POR 2 CHECADAS
                                    else if (Lunes.Count == 3)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Lunes[1].ToString()) != null) && (Convert.ToDateTime(Lunes[1].ToString()) >= SalidaTarde && Convert.ToDateTime(Lunes[1].ToString()) < SalidaTarde.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[1].ToString()), SalidaTarde);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Lunes[1].ToString()) != null) && (Convert.ToDateTime(Lunes[1].ToString()) < SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[1].ToString()), SalidaTarde);
                                        }
                                    }//IF POR 2 CHECADAS
                                    else if (Lunes.Count == 4)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Lunes[3].ToString()) != null) && (Convert.ToDateTime(Lunes[3].ToString()) >= SalidaTarde && Convert.ToDateTime(Lunes[3].ToString()) < SalidaTarde.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[1].ToString()), SalidaTarde);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Lunes[3].ToString()) != null) && (Convert.ToDateTime(Lunes[3].ToString()) < SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[3].ToString()), SalidaTarde);
                                        }
                                    }//IF POR 4 CHECADAS
                                }//FIN TURNO TARDE
                                 //////////////////////////// - TDIA - TDIA - TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA"))//TURNO NOCHE
                                {
                                    TimeSpan Tdia = new TimeSpan(15, 30, 00);//Variable con el horario 19:00 hrs
                                    SalidaTDIA = Convert.ToDateTime(Lunes[0].ToString()).Date + Tdia;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Lunes ----------------------------------------------------------------------//
                                    if (Lunes.Count == 2)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Lunes[1].ToString()) != null) && (Convert.ToDateTime(Lunes[1].ToString()) >= SalidaTDIA && Convert.ToDateTime(Lunes[1].ToString()) < SalidaTDIA.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[1].ToString()), SalidaTDIA);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Lunes[1].ToString()) != null) && (Convert.ToDateTime(Lunes[1].ToString()) < SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[1].ToString()), SalidaTDIA);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Lunes[1].ToString()) != null) && (Convert.ToDateTime(Lunes[1].ToString()) > SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[1].ToString()), SalidaTDIA);
                                        }
                                    }
                                    if (Lunes.Count >= 4)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Lunes[3].ToString()) != null) && (Convert.ToDateTime(Lunes[3].ToString()) >= SalidaTDIA && Convert.ToDateTime(Lunes[3].ToString()) < SalidaTDIA.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[3].ToString()), SalidaTDIA);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Lunes[3].ToString()) != null) && (Convert.ToDateTime(Lunes[3].ToString()) < SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[3].ToString()), SalidaTDIA);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Lunes[3].ToString()) != null) && (Convert.ToDateTime(Lunes[3].ToString()) > SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[3].ToString()), SalidaTDIA);
                                        }
                                    }

                                }//Fin noche
                                 ////////////////////////// - TFIJ - TFIJ - TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ /////////////////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ"))//TURNO TFIJ
                                {
                                    TimeSpan Tfij = new TimeSpan(17, 00, 00);//Variable con el horario 19:00 hrs
                                    SalidaTFIJ = Convert.ToDateTime(Lunes[0].ToString()).Date + Tfij;//Establece fecha de la checada con la hora de ts
                                                                                                     //----------------------------------------------------------- Lunes ----------------------------------------------------------------------//
                                    if (Lunes.Count == 2)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Lunes[1].ToString()) != null) && (Convert.ToDateTime(Lunes[1].ToString()) > SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[1].ToString()), SalidaTFIJ);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Lunes[1].ToString()) != null) && (Convert.ToDateTime(Lunes[1].ToString()) < SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[1].ToString()), SalidaTFIJ);
                                        }
                                    }
                                    if (Lunes.Count >= 4)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Lunes[3].ToString()) != null) && (Convert.ToDateTime(Lunes[3].ToString()) >= SalidaTFIJ && Convert.ToDateTime(Lunes[3].ToString()) < SalidaTFIJ.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[3].ToString()), SalidaTFIJ);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Lunes[3].ToString()) != null) && (Convert.ToDateTime(Lunes[3].ToString()) < SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[3].ToString()), SalidaTFIJ);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Lunes[3].ToString()) != null) && (Convert.ToDateTime(Lunes[3].ToString()) > SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[3].ToString()), SalidaTFIJ);
                                        }
                                    }
                                }//Fin TFIJ
                                 ///////////////////////// - TNO2 - TNO2 - TNO2 - TNO2 - TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2 - TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 ///////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2"))//TURNO NOCHE
                                {
                                    TimeSpan Tno2 = new TimeSpan(19, 00, 00);//Variable con el horario 19:00 hrs
                                    SalidaTNO2 = Convert.ToDateTime(Lunes[0].ToString()).Date + Tno2;//Establece fecha de la checada con la hora de ts
                                                                                                     //----------------------------------------------------------- Lunes ----------------------------------------------------------------------//
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Lunes[3].ToString()) != null) && (Convert.ToDateTime(Lunes[3].ToString()) >= SalidaTNO2 && Convert.ToDateTime(Lunes[3].ToString()) < SalidaTNO2.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[3].ToString()), SalidaTNO2);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Lunes[3].ToString()) != null) && (Convert.ToDateTime(Lunes[3].ToString()) < SalidaTNO2))
                                    {
                                        c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[3].ToString()), SalidaTNO2);
                                    }
                                }//Fin TNO2
                            } catch (Exception ex) { }//TRY DIA LUNES
                            break;

                        case "Martes":
                        case "Tuesday":
                            try
                            {
                                ///////////// - NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE /////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE"))//TURNO NOCHE
                                {
                                    TimeSpan Noche = new TimeSpan(07, 00, 00);//Variable con el horario 19:00 hrs
                                    SalidaNoche = Convert.ToDateTime(Martes[0].ToString()).Date + Noche;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Martes ----------------------------------------------------------------------//
                                    if (Martes.Count == 3)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Martes[0].ToString()) != null) && (Convert.ToDateTime(Martes[0].ToString()) >= SalidaNoche && Convert.ToDateTime(Martes[0].ToString()) < SalidaNoche.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), SalidaNoche);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Martes[0].ToString()) != null) && (Convert.ToDateTime(Martes[0].ToString()) < SalidaNoche))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), SalidaNoche);
                                        }
                                    }
                                    else if (Martes.Count == 2)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Martes[0].ToString()) != null) && (Convert.ToDateTime(Martes[0].ToString()) >= SalidaNoche && Convert.ToDateTime(Martes[0].ToString()) < SalidaNoche.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), SalidaNoche);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Martes[0].ToString()) != null) && (Convert.ToDateTime(Martes[0].ToString()) < SalidaNoche))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), SalidaNoche);
                                        }
                                    }
                                }//Fin noche
                                //////////// - TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE///////////////////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE"))//TURNO NOCHE
                                {
                                    TimeSpan Tarde = new TimeSpan(23, 00, 00);//Variable con el horario 19:00 hrs
                                    DateTime SalidaTarde = Convert.ToDateTime(Martes[0].ToString()).Date + Tarde;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Martes ----------------------------------------------------------------------//
                                    if (Martes.Count == 2)//En caso de 2 checadas
                                    {
                                        if (Convert.ToDateTime(Martes[1].ToString()) >= SalidaTarde && Convert.ToDateTime(Martes[1].ToString()) < SalidaTarde.AddMinutes(30))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[1].ToString()), SalidaTarde);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Martes[1].ToString()) != null) && (Convert.ToDateTime(Martes[1].ToString()) < SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), SalidaTarde);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Martes[1].ToString()) != null) && (Convert.ToDateTime(Martes[1].ToString()) > SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[0].ToString()), SalidaTarde);
                                        }
                                    }//IF POR 2 CHECADAS
                                    else if (Martes.Count == 3)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Martes[1].ToString()) != null) && (Convert.ToDateTime(Martes[1].ToString()) >= SalidaTarde && Convert.ToDateTime(Martes[1].ToString()) < SalidaTarde.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[1].ToString()), SalidaTarde);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Martes[1].ToString()) != null) && (Convert.ToDateTime(Martes[1].ToString()) < SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[1].ToString()), SalidaTarde);
                                        }
                                    }//IF POR 2 CHECADAS
                                    else if (Martes.Count == 4)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Martes[3].ToString()) != null) && (Convert.ToDateTime(Martes[3].ToString()) >= SalidaTarde && Convert.ToDateTime(Martes[3].ToString()) < SalidaTarde.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[1].ToString()), SalidaTarde);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Martes[3].ToString()) != null) && (Convert.ToDateTime(Martes[3].ToString()) < SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[3].ToString()), SalidaTarde);
                                        }
                                    }//IF POR 4 CHECADAS
                                }//FIN TURNO TARDE
                                 //////////////////////////// - TDIA - TDIA - TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA"))//TURNO NOCHE
                                {
                                    TimeSpan Tdia = new TimeSpan(15, 30, 00);//Variable con el horario 19:00 hrs
                                    SalidaTDIA = Convert.ToDateTime(Martes[0].ToString()).Date + Tdia;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Martes ----------------------------------------------------------------------//
                                    if (Martes.Count == 2)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Martes[1].ToString()) != null) && (Convert.ToDateTime(Martes[1].ToString()) >= SalidaTDIA && Convert.ToDateTime(Martes[1].ToString()) < SalidaTDIA.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[1].ToString()), SalidaTDIA);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Martes[1].ToString()) != null) && (Convert.ToDateTime(Martes[1].ToString()) < SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[1].ToString()), SalidaTDIA);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Martes[1].ToString()) != null) && (Convert.ToDateTime(Martes[1].ToString()) > SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[1].ToString()), SalidaTDIA);
                                        }
                                    }
                                    if (Martes.Count >= 4)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Martes[3].ToString()) != null) && (Convert.ToDateTime(Martes[3].ToString()) >= SalidaTDIA && Convert.ToDateTime(Martes[3].ToString()) < SalidaTDIA.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[3].ToString()), SalidaTDIA);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Martes[3].ToString()) != null) && (Convert.ToDateTime(Martes[3].ToString()) < SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[3].ToString()), SalidaTDIA);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Martes[3].ToString()) != null) && (Convert.ToDateTime(Martes[3].ToString()) > SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[3].ToString()), SalidaTDIA);
                                        }
                                    }
                                }//Fin noche
                                 ////////////////////////// - TFIJ - TFIJ - TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ /////////////////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ"))//TURNO NOCHE
                                {
                                    TimeSpan Tfij = new TimeSpan(17, 00, 00);//Variable con el horario 19:00 hrs
                                    SalidaTFIJ = Convert.ToDateTime(Martes[0].ToString()).Date + Tfij;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Martes ----------------------------------------------------------------------//
                                    if (Martes.Count == 2)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Martes[1].ToString()) != null) && (Convert.ToDateTime(Martes[1].ToString()) > SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[1].ToString()), SalidaTFIJ);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Martes[1].ToString()) != null) && (Convert.ToDateTime(Martes[1].ToString()) < SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[1].ToString()), SalidaTFIJ);
                                        }
                                    }
                                    if (Martes.Count >= 4)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Martes[3].ToString()) != null) && (Convert.ToDateTime(Martes[3].ToString()) >= SalidaTFIJ && Convert.ToDateTime(Martes[3].ToString()) < SalidaTFIJ.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[3].ToString()), SalidaTFIJ);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Martes[3].ToString()) != null) && (Convert.ToDateTime(Martes[3].ToString()) < SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[3].ToString()), SalidaTFIJ);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Martes[3].ToString()) != null) && (Convert.ToDateTime(Martes[3].ToString()) > SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[3].ToString()), SalidaTFIJ);
                                        }
                                    }
                                }//Fin TFIJ
                                 ///////////////////////// - TNO2 - TNO2 - TNO2 - TNO2 - TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2 - TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 ///////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2"))//TURNO NOCHE
                                {
                                    TimeSpan Tno2 = new TimeSpan(19, 00, 00);//Variable con el horario 19:00 hrs
                                    SalidaTNO2 = Convert.ToDateTime(Martes[0].ToString()).Date + Tno2;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Martes ----------------------------------------------------------------------//
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Martes[3].ToString()) != null) && (Convert.ToDateTime(Martes[3].ToString()) >= SalidaTNO2 && Convert.ToDateTime(Martes[3].ToString()) < SalidaTNO2.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[3].ToString()), SalidaTNO2);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Martes[3].ToString()) != null) && (Convert.ToDateTime(Martes[3].ToString()) < SalidaTNO2))
                                    {
                                        c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[3].ToString()), SalidaTNO2);
                                    }
                                }//Fin TNO2
                            } catch (Exception ex) { }//TRY DIA Martes
                            break;
                        case "Miercoles":
                        case "Wednesday":
                            try
                            {
                                ///////////// - NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE /////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE"))//TURNO NOCHE
                                {
                                    TimeSpan Noche = new TimeSpan(07, 00, 00);//Variable con el horario 19:00 hrs
                                    SalidaNoche = Convert.ToDateTime(Miercoles[0].ToString()).Date + Noche;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Martes ----------------------------------------------------------------------//
                                    if (Miercoles.Count == 3)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Miercoles[0].ToString()) != null) && (Convert.ToDateTime(Miercoles[0].ToString()) >= SalidaNoche && Convert.ToDateTime(Miercoles[0].ToString()) < SalidaNoche.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), SalidaNoche);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Miercoles[0].ToString()) != null) && (Convert.ToDateTime(Miercoles[0].ToString()) < SalidaNoche))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), SalidaNoche);
                                        }
                                    }
                                    else if (Miercoles.Count == 2)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Miercoles[0].ToString()) != null) && (Convert.ToDateTime(Miercoles[0].ToString()) >= SalidaNoche && Convert.ToDateTime(Miercoles[0].ToString()) < SalidaNoche.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), SalidaNoche);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Miercoles[0].ToString()) != null) && (Convert.ToDateTime(Miercoles[0].ToString()) < SalidaNoche))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[0].ToString()), SalidaNoche);
                                        }
                                    }
                                }//Fin noche
                                //////////// - TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE///////////////////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE"))//TURNO NOCHE
                                {
                                    TimeSpan Tarde = new TimeSpan(23, 00, 00);//Variable con el horario 19:00 hrs
                                    DateTime SalidaTarde = Convert.ToDateTime(Miercoles[0].ToString()).Date + Tarde;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Miercoles ----------------------------------------------------------------------//
                                    if (Miercoles.Count == 2)//En caso de 2 checadas
                                    {
                                        if (Convert.ToDateTime(Miercoles[1].ToString()) >= SalidaTarde && Convert.ToDateTime(Miercoles[1].ToString()) < SalidaTarde.AddMinutes(30))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[1].ToString()), SalidaTarde);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Miercoles[1].ToString()) != null) && (Convert.ToDateTime(Miercoles[1].ToString()) < SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[1].ToString()), SalidaTarde);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Miercoles[1].ToString()) != null) && (Convert.ToDateTime(Miercoles[1].ToString()) > SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[1].ToString()), SalidaTarde);
                                        }
                                    }//IF POR 2 CHECADAS
                                    else if (Miercoles.Count == 3)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Miercoles[1].ToString()) != null) && (Convert.ToDateTime(Miercoles[1].ToString()) >= SalidaTarde && Convert.ToDateTime(Miercoles[1].ToString()) < SalidaTarde.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[1].ToString()), SalidaTarde);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Miercoles[1].ToString()) != null) && (Convert.ToDateTime(Miercoles[1].ToString()) < SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[1].ToString()), SalidaTarde);
                                        }
                                    }//IF POR 2 CHECADAS
                                    else if (Miercoles.Count == 4)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Miercoles[3].ToString()) != null) && (Convert.ToDateTime(Miercoles[3].ToString()) >= SalidaTarde && Convert.ToDateTime(Miercoles[3].ToString()) < SalidaTarde.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[1].ToString()), SalidaTarde);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Miercoles[3].ToString()) != null) && (Convert.ToDateTime(Miercoles[3].ToString()) < SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[3].ToString()), SalidaTarde);
                                        }
                                    }//IF POR 4 CHECADAS
                                }//FIN TURNO TARDE
                                 //////////////////////////// - TDIA - TDIA - TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA"))//TURNO NOCHE
                                {
                                    TimeSpan Tdia = new TimeSpan(15, 30, 00);//Variable con el horario 19:00 hrs
                                    SalidaTDIA = Convert.ToDateTime(Miercoles[0].ToString()).Date + Tdia;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Miercoles ----------------------------------------------------------------------//
                                    if (Miercoles.Count == 2)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Miercoles[3].ToString()) != null) && (Convert.ToDateTime(Miercoles[1].ToString()) >= SalidaTDIA && Convert.ToDateTime(Miercoles[1].ToString()) < SalidaTDIA.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[3].ToString()), SalidaTDIA);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Miercoles[1].ToString()) != null) && (Convert.ToDateTime(Miercoles[1].ToString()) < SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[3].ToString()), SalidaTDIA);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Miercoles[1].ToString()) != null) && (Convert.ToDateTime(Miercoles[1].ToString()) > SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[3].ToString()), SalidaTDIA);
                                        }
                                    }
                                    if (Miercoles.Count >= 4)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Miercoles[3].ToString()) != null) && (Convert.ToDateTime(Miercoles[3].ToString()) >= SalidaTDIA && Convert.ToDateTime(Miercoles[3].ToString()) < SalidaTDIA.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[3].ToString()), SalidaTDIA);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Miercoles[3].ToString()) != null) && (Convert.ToDateTime(Miercoles[3].ToString()) < SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[3].ToString()), SalidaTDIA);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Miercoles[3].ToString()) != null) && (Convert.ToDateTime(Miercoles[3].ToString()) > SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[3].ToString()), SalidaTDIA);
                                        }
                                    }
                                }//Fin noche
                                 ////////////////////////// - TFIJ - TFIJ - TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ /////////////////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ"))//TURNO NOCHE
                                {
                                    TimeSpan Tfij = new TimeSpan(17, 00, 00);//Variable con el horario 19:00 hrs
                                    SalidaTFIJ = Convert.ToDateTime(Miercoles[0].ToString()).Date + Tfij;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Miercoles ----------------------------------------------------------------------//
                                    if (Miercoles.Count == 2)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Miercoles[1].ToString()) != null) && (Convert.ToDateTime(Miercoles[1].ToString()) > SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[1].ToString()), SalidaTFIJ);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Miercoles[1].ToString()) != null) && (Convert.ToDateTime(Miercoles[1].ToString()) < SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[1].ToString()), SalidaTFIJ);
                                        }
                                    }
                                    if (Miercoles.Count >= 4)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Miercoles[3].ToString()) != null) && (Convert.ToDateTime(Miercoles[3].ToString()) >= SalidaTFIJ && Convert.ToDateTime(Miercoles[3].ToString()) < SalidaTFIJ.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[3].ToString()), SalidaTFIJ);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Miercoles[3].ToString()) != null) && (Convert.ToDateTime(Miercoles[3].ToString()) < SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[3].ToString()), SalidaTFIJ);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Miercoles[3].ToString()) != null) && (Convert.ToDateTime(Miercoles[3].ToString()) > SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[3].ToString()), SalidaTFIJ);
                                        }
                                    }
                                }//Fin TFIJ
                                 ///////////////////////// - TNO2 - TNO2 - TNO2 - TNO2 - TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2 - TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 ///////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2"))//TURNO NOCHE
                                {
                                    TimeSpan Tno2 = new TimeSpan(19, 00, 00);//Variable con el horario 19:00 hrs
                                    SalidaTNO2 = Convert.ToDateTime(Miercoles[0].ToString()).Date + Tno2;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Miercoles ----------------------------------------------------------------------//
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Miercoles[3].ToString()) != null) && (Convert.ToDateTime(Miercoles[3].ToString()) >= SalidaTNO2 && Convert.ToDateTime(Miercoles[3].ToString()) < SalidaTNO2.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[3].ToString()), SalidaTNO2);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Miercoles[3].ToString()) != null) && (Convert.ToDateTime(Miercoles[3].ToString()) < SalidaTNO2))
                                    {
                                        c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[3].ToString()), SalidaTNO2);
                                    }
                                }//Fin TNO2
                            }
                            catch (Exception ex) { }//TRY DIA Miercoles
                            break;
                        case "Jueves":
                        case "Thursday":
                            try
                            {
                                ///////////// - NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE /////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE"))//TURNO NOCHE
                                {
                                    TimeSpan Noche = new TimeSpan(07, 00, 00);//Variable con el horario 19:00 hrs
                                    SalidaNoche = Convert.ToDateTime(Jueves[0].ToString()).Date + Noche;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Martes ----------------------------------------------------------------------//
                                    if (Jueves.Count == 3)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Jueves[0].ToString()) != null) && (Convert.ToDateTime(Jueves[0].ToString()) >= SalidaNoche && Convert.ToDateTime(Jueves[0].ToString()) < SalidaNoche.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), SalidaNoche);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Jueves[0].ToString()) != null) && (Convert.ToDateTime(Jueves[0].ToString()) < SalidaNoche))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), SalidaNoche);
                                        }
                                    }
                                    else if (Jueves.Count == 2)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Jueves[0].ToString()) != null) && (Convert.ToDateTime(Jueves[0].ToString()) >= SalidaNoche && Convert.ToDateTime(Jueves[0].ToString()) < SalidaNoche.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), SalidaNoche);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Jueves[0].ToString()) != null) && (Convert.ToDateTime(Jueves[0].ToString()) < SalidaNoche))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[0].ToString()), SalidaNoche);
                                        }
                                    }
                                }//Fin noche
                                //////////// - TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE///////////////////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE"))//TURNO NOCHE
                                {
                                    TimeSpan Tarde = new TimeSpan(23, 00, 00);//Variable con el horario 19:00 hrs
                                    DateTime SalidaTarde = Convert.ToDateTime(Jueves[0].ToString()).Date + Tarde;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Jueves ----------------------------------------------------------------------//
                                    if (Jueves.Count == 2)//En caso de 2 checadas
                                    {
                                        if (Convert.ToDateTime(Jueves[1].ToString()) >= SalidaTarde && Convert.ToDateTime(Jueves[1].ToString()) < SalidaTarde.AddMinutes(30))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[1].ToString()), SalidaTarde);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Jueves[1].ToString()) != null) && (Convert.ToDateTime(Jueves[1].ToString()) < SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[1].ToString()), SalidaTarde);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Jueves[1].ToString()) != null) && (Convert.ToDateTime(Jueves[1].ToString()) > SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[1].ToString()), SalidaTarde);
                                        }
                                    }//IF POR 2 CHECADAS
                                    else if (Jueves.Count == 3)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Jueves[1].ToString()) != null) && (Convert.ToDateTime(Jueves[1].ToString()) >= SalidaTarde && Convert.ToDateTime(Jueves[1].ToString()) < SalidaTarde.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[1].ToString()), SalidaTarde);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Jueves[1].ToString()) != null) && (Convert.ToDateTime(Jueves[1].ToString()) < SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[1].ToString()), SalidaTarde);
                                        }
                                    }//IF POR 2 CHECADAS
                                    else if (Jueves.Count == 4)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Jueves[3].ToString()) != null) && (Convert.ToDateTime(Jueves[3].ToString()) >= SalidaTarde && Convert.ToDateTime(Jueves[3].ToString()) < SalidaTarde.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[1].ToString()), SalidaTarde);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Jueves[3].ToString()) != null) && (Convert.ToDateTime(Jueves[3].ToString()) < SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[3].ToString()), SalidaTarde);
                                        }
                                    }//IF POR 4 CHECADAS
                                }//FIN TURNO TARDE
                                 //////////////////////////// - TDIA - TDIA - TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA"))//TURNO NOCHE
                                {
                                    TimeSpan Tdia = new TimeSpan(15, 30, 00);//Variable con el horario 19:00 hrs
                                    SalidaTDIA = Convert.ToDateTime(Jueves[0].ToString()).Date + Tdia;//Establece fecha de la checada con la hora de ts
                                                                                                      //----------------------------------------------------------- Jueves ----------------------------------------------------------------------//
                                    if (Jueves.Count == 2)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Jueves[1].ToString()) != null) && (Convert.ToDateTime(Jueves[1].ToString()) >= SalidaTDIA && Convert.ToDateTime(Jueves[1].ToString()) < SalidaTDIA.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[1].ToString()), SalidaTDIA);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Jueves[1].ToString()) != null) && (Convert.ToDateTime(Jueves[1].ToString()) < SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[1].ToString()), SalidaTDIA);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Jueves[1].ToString()) != null) && (Convert.ToDateTime(Jueves[1].ToString()) > SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[1].ToString()), SalidaTDIA);
                                        }
                                    }
                                    if (Jueves.Count >= 4)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Jueves[3].ToString()) != null) && (Convert.ToDateTime(Jueves[3].ToString()) >= SalidaTDIA && Convert.ToDateTime(Jueves[3].ToString()) < SalidaTDIA.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[3].ToString()), SalidaTDIA);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Jueves[3].ToString()) != null) && (Convert.ToDateTime(Jueves[3].ToString()) < SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[3].ToString()), SalidaTDIA);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Jueves[3].ToString()) != null) && (Convert.ToDateTime(Jueves[3].ToString()) > SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[3].ToString()), SalidaTDIA);
                                        }
                                    }
                                }//Fin noche
                                 ////////////////////////// - TFIJ - TFIJ - TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ /////////////////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ"))//TURNO TFIJ
                                {
                                    TimeSpan Tfij = new TimeSpan(17, 00, 00);//Variable con el horario 19:00 hrs
                                    SalidaTFIJ = Convert.ToDateTime(Jueves[0].ToString()).Date + Tfij;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Jueves ----------------------------------------------------------------------//
                                    if (Jueves.Count == 2)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Jueves[1].ToString()) != null) && (Convert.ToDateTime(Jueves[1].ToString()) > SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[1].ToString()), SalidaTFIJ);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Jueves[1].ToString()) != null) && (Convert.ToDateTime(Jueves[1].ToString()) < SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[1].ToString()), SalidaTFIJ);
                                        }
                                    }
                                    if (Jueves.Count >= 4)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Jueves[3].ToString()) != null) && (Convert.ToDateTime(Jueves[3].ToString()) >= SalidaTFIJ && Convert.ToDateTime(Jueves[3].ToString()) < SalidaTFIJ.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[3].ToString()), SalidaTFIJ);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Jueves[3].ToString()) != null) && (Convert.ToDateTime(Jueves[3].ToString()) < SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[3].ToString()), SalidaTFIJ);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Jueves[3].ToString()) != null) && (Convert.ToDateTime(Jueves[3].ToString()) > SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[3].ToString()), SalidaTFIJ);
                                        }
                                    }
                                }//Fin TFIJ
                                 ///////////////////////// - TNO2 - TNO2 - TNO2 - TNO2 - TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2 - TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 ///////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2"))//TURNO NOCHE
                                {
                                    TimeSpan Tno2 = new TimeSpan(19, 00, 00);//Variable con el horario 19:00 hrs
                                    SalidaTNO2 = Convert.ToDateTime(Jueves[0].ToString()).Date + Tno2;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Jueves ----------------------------------------------------------------------//
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Jueves[3].ToString()) != null) && (Convert.ToDateTime(Jueves[3].ToString()) >= SalidaTNO2 && Convert.ToDateTime(Jueves[3].ToString()) < SalidaTNO2.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[3].ToString()), SalidaTNO2);//Actualiza la checada al turno asignado
                                    }

                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Jueves[3].ToString()) != null) && (Convert.ToDateTime(Jueves[3].ToString()) < SalidaTNO2))
                                    {
                                        c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[3].ToString()), SalidaTNO2);
                                    }
                                }//Fin TNO2
                            }
                            catch (Exception ex) { }//TRY DIA Martes
                            break;

                        case "Viernes":
                        case "Friday":
                            try
                            {
                                ///////////// - NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE /////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE"))//TURNO NOCHE
                                {
                                    TimeSpan Noche = new TimeSpan(07, 00, 00);//Variable con el horario 19:00 hrs
                                    SalidaNoche = Convert.ToDateTime(Viernes[0].ToString()).Date + Noche;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Martes ----------------------------------------------------------------------//
                                    if (Viernes.Count == 3)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Viernes[0].ToString()) != null) && (Convert.ToDateTime(Viernes[0].ToString()) >= SalidaNoche && Convert.ToDateTime(Viernes[0].ToString()) < SalidaNoche.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), SalidaNoche);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Viernes[0].ToString()) != null) && (Convert.ToDateTime(Viernes[0].ToString()) < SalidaNoche))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), SalidaNoche);
                                        }
                                    }
                                    else if (Viernes.Count == 2)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Viernes[0].ToString()) != null) && (Convert.ToDateTime(Viernes[0].ToString()) >= SalidaNoche && Convert.ToDateTime(Viernes[0].ToString()) < SalidaNoche.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), SalidaNoche);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Viernes[0].ToString()) != null) && (Convert.ToDateTime(Viernes[0].ToString()) < SalidaNoche))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[0].ToString()), SalidaNoche);
                                        }
                                    }
                                }//Fin noche
                                //////////// - TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE///////////////////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE"))//TURNO NOCHE
                                {
                                    TimeSpan Tarde = new TimeSpan(23, 00, 00);//Variable con el horario 19:00 hrs
                                    DateTime SalidaTarde = Convert.ToDateTime(Viernes[0].ToString()).Date + Tarde;//Establece fecha de la checada con la hora de ts
                                                                                                                  //----------------------------------------------------------- Viernes ----------------------------------------------------------------------//
                                    if (Viernes.Count == 2)//En caso de 2 checadas
                                    {
                                        if (Convert.ToDateTime(Viernes[1].ToString()) >= SalidaTarde && Convert.ToDateTime(Viernes[1].ToString()) < SalidaTarde.AddMinutes(30))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[1].ToString()), SalidaTarde);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Viernes[1].ToString()) != null) && (Convert.ToDateTime(Viernes[1].ToString()) < SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[1].ToString()), SalidaTarde);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Viernes[1].ToString()) != null) && (Convert.ToDateTime(Viernes[1].ToString()) > SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[1].ToString()), SalidaTarde);
                                        }
                                    }//IF POR 2 CHECADAS
                                    else if (Viernes.Count == 3)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Viernes[1].ToString()) != null) && (Convert.ToDateTime(Viernes[1].ToString()) >= SalidaTarde && Convert.ToDateTime(Viernes[1].ToString()) < SalidaTarde.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[1].ToString()), SalidaTarde);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Viernes[1].ToString()) != null) && (Convert.ToDateTime(Viernes[1].ToString()) < SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[1].ToString()), SalidaTarde);
                                        }
                                    }//IF POR 2 CHECADAS
                                    else if (Viernes.Count == 4)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Viernes[3].ToString()) != null) && (Convert.ToDateTime(Viernes[3].ToString()) >= SalidaTarde && Convert.ToDateTime(Viernes[3].ToString()) < SalidaTarde.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[1].ToString()), SalidaTarde);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Viernes[3].ToString()) != null) && (Convert.ToDateTime(Viernes[3].ToString()) < SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[3].ToString()), SalidaTarde);
                                        }
                                    }//IF POR 4 CHECADAS
                                }//FIN TURNO TARDE
                                 //////////////////////////// - TDIA - TDIA - TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA"))//TURNO NOCHE
                                {
                                    TimeSpan Tdia = new TimeSpan(15, 30, 00);//Variable con el horario 19:00 hrs
                                    SalidaTDIA = Convert.ToDateTime(Viernes[0].ToString()).Date + Tdia;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Viernes ----------------------------------------------------------------------//
                                    if (Viernes.Count == 2)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Viernes[1].ToString()) != null) && (Convert.ToDateTime(Viernes[1].ToString()) >= SalidaTDIA && Convert.ToDateTime(Viernes[1].ToString()) < SalidaTDIA.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[1].ToString()), SalidaTDIA);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Viernes[1].ToString()) != null) && (Convert.ToDateTime(Viernes[1].ToString()) < SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[1].ToString()), SalidaTDIA);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Viernes[1].ToString()) != null) && (Convert.ToDateTime(Viernes[1].ToString()) > SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[1].ToString()), SalidaTDIA);
                                        }
                                    }
                                    if (Viernes.Count >= 4)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Viernes[3].ToString()) != null) && (Convert.ToDateTime(Viernes[3].ToString()) >= SalidaTDIA && Convert.ToDateTime(Viernes[3].ToString()) < SalidaTDIA.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[3].ToString()), SalidaTDIA);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Viernes[3].ToString()) != null) && (Convert.ToDateTime(Viernes[3].ToString()) < SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[3].ToString()), SalidaTDIA);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Viernes[3].ToString()) != null) && (Convert.ToDateTime(Viernes[3].ToString()) > SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[3].ToString()), SalidaTDIA);
                                        }
                                    }
                                }//Fin noche
                                 ////////////////////////// - TFIJ - TFIJ - TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ /////////////////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ"))//TURNO NOCHE
                                {
                                    TimeSpan Tfij = new TimeSpan(17, 00, 00);//Variable con el horario 19:00 hrs
                                    SalidaTFIJ = Convert.ToDateTime(Viernes[0].ToString()).Date + Tfij;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Viernes ----------------------------------------------------------------------//
                                    if (Viernes.Count == 2)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Viernes[1].ToString()) != null) && (Convert.ToDateTime(Viernes[1].ToString()) > SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[1].ToString()), SalidaTFIJ);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Viernes[1].ToString()) != null) && (Convert.ToDateTime(Viernes[1].ToString()) < SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[1].ToString()), SalidaTFIJ);
                                        }
                                    }
                                    if (Viernes.Count >= 4)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Viernes[3].ToString()) != null) && (Convert.ToDateTime(Viernes[3].ToString()) >= SalidaTFIJ && Convert.ToDateTime(Viernes[3].ToString()) < SalidaTFIJ.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[3].ToString()), SalidaTFIJ);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Viernes[3].ToString()) != null) && (Convert.ToDateTime(Viernes[3].ToString()) < SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[3].ToString()), SalidaTFIJ);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Viernes[3].ToString()) != null) && (Convert.ToDateTime(Viernes[3].ToString()) > SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[3].ToString()), SalidaTFIJ);
                                        }
                                    }
                                }//Fin TFIJ
                                 ///////////////////////// - TNO2 - TNO2 - TNO2 - TNO2 - TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2 - TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 ///////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2"))//TURNO NOCHE
                                {
                                    TimeSpan Tno2 = new TimeSpan(19, 00, 00);//Variable con el horario 19:00 hrs
                                    SalidaTNO2 = Convert.ToDateTime(Viernes[0].ToString()).Date + Tno2;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Viernes ----------------------------------------------------------------------//
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Viernes[3].ToString()) != null) && (Convert.ToDateTime(Viernes[3].ToString()) >= SalidaTNO2 && Convert.ToDateTime(Viernes[3].ToString()) < SalidaTNO2.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[3].ToString()), SalidaTNO2);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Viernes[3].ToString()) != null) && (Convert.ToDateTime(Viernes[3].ToString()) < SalidaTNO2))
                                    {
                                        c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[3].ToString()), SalidaTNO2);
                                    }
                                }//Fin TNO2
                            }
                            catch (Exception ex) { }//TRY DIA Martes
                            break;

                        case "Sábado":
                        case "Saturday":
                            try
                            {
                                ///////////// - NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE- NOCHE - NOCHE /////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE"))//TURNO NOCHE
                                {
                                    TimeSpan Noche = new TimeSpan(07, 00, 00);//Variable con el horario 19:00 hrs
                                    SalidaNoche = Convert.ToDateTime(Sabado[0].ToString()).Date + Noche;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Sabado ----------------------------------------------------------------------//
                                    if (Sabado.Count == 3)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Sabado[0].ToString()) != null) && (Convert.ToDateTime(Sabado[0].ToString()) >= SalidaNoche && Convert.ToDateTime(Sabado[0].ToString()) < SalidaNoche.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), SalidaNoche);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Sabado[0].ToString()) != null) && (Convert.ToDateTime(Sabado[0].ToString()) < SalidaNoche))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), SalidaNoche);
                                        }
                                    }
                                    if (Sabado.Count == 1)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Sabado[0].ToString()) != null) && (Convert.ToDateTime(Sabado[0].ToString()) >= SalidaNoche && Convert.ToDateTime(Sabado[0].ToString()) < SalidaNoche.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), SalidaNoche);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("NOCHE") && (Convert.ToDateTime(Sabado[0].ToString()) != null) && (Convert.ToDateTime(Sabado[0].ToString()) < SalidaNoche))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[0].ToString()), SalidaNoche);
                                        }
                                    }
                                }//Fin noche
                                 ///////////// -SAB - SAB -SAB - SAB -SAB - SAB -SAB - SAB -SAB - SAB -SAB - SAB -SAB - SAB -SAB - SAB -SAB - SAB -SAB - SAB -SAB - SAB -SAB - SAB  -SAB - SAB  -SAB - SAB  -SAB - SAB  -SAB - SAB  -SAB - SAB  -SAB - SAB  -SAB - SAB  -SAB - SAB  -SAB - SAB  -SAB - SAB /////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB"))//TURNO NOCHE
                                {
                                    TimeSpan SAB = new TimeSpan(13, 30, 00);//Variable con el horario 19:00 hrs
                                    SalidaSAB = Convert.ToDateTime(Sabado[0].ToString()).Date + SAB;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Sabado ----------------------------------------------------------------------//
                                    if (Sabado.Count == 2)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Sabado[1].ToString()) != null) && (Convert.ToDateTime(Sabado[1].ToString()) >= SalidaSAB && Convert.ToDateTime(Sabado[1].ToString()) < SalidaSAB.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[1].ToString()), SalidaSAB);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Sabado[1].ToString()) != null) && (Convert.ToDateTime(Sabado[1].ToString()) > SalidaSAB))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[1].ToString()), SalidaSAB);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("SAB") && (Convert.ToDateTime(Sabado[1].ToString()) != null) && (Convert.ToDateTime(Sabado[1].ToString()) < SalidaSAB))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[1].ToString()), SalidaSAB);
                                        }
                                    }
                                }//Fin noche
                                //////////// - TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE- TARDE - TARDE///////////////////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE"))//TURNO NOCHE
                                {
                                    TimeSpan Tarde = new TimeSpan(23, 00, 00);//Variable con el horario 19:00 hrs
                                    DateTime SalidaTarde = Convert.ToDateTime(Sabado[0].ToString()).Date + Tarde;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Sabado ----------------------------------------------------------------------//
                                    if (Sabado.Count == 2)//En caso de 2 checadas
                                    {
                                        if (Convert.ToDateTime(Sabado[1].ToString()) >= SalidaTarde && Convert.ToDateTime(Sabado[1].ToString()) < SalidaTarde.AddMinutes(30))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[1].ToString()), SalidaTarde);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Sabado[1].ToString()) != null) && (Convert.ToDateTime(Sabado[1].ToString()) < SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[1].ToString()), SalidaTarde);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Sabado[1].ToString()) != null) && (Convert.ToDateTime(Sabado[1].ToString()) > SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[1].ToString()), SalidaTarde);
                                        }
                                    }//IF POR 2 CHECADAS
                                    else if (Sabado.Count == 3)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Sabado[1].ToString()) != null) && (Convert.ToDateTime(Sabado[1].ToString()) >= SalidaTarde && Convert.ToDateTime(Sabado[1].ToString()) < SalidaTarde.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[1].ToString()), SalidaTarde);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Sabado[1].ToString()) != null) && (Convert.ToDateTime(Sabado[1].ToString()) < SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[1].ToString()), SalidaTarde);
                                        }
                                    }//IF POR 2 CHECADAS
                                    else if (Sabado.Count == 4)//En caso de 3 checadas
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Sabado[3].ToString()) != null) && (Convert.ToDateTime(Sabado[3].ToString()) >= SalidaTarde && Convert.ToDateTime(Sabado[3].ToString()) < SalidaTarde.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[1].ToString()), SalidaTarde);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TARDE") && (Convert.ToDateTime(Sabado[3].ToString()) != null) && (Convert.ToDateTime(Sabado[3].ToString()) < SalidaTarde))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[3].ToString()), SalidaTarde);
                                        }
                                    }//IF POR 4 CHECADAS
                                }//FIN TURNO TARDE
                                 //////////////////////////// - TDIA - TDIA - TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA- TDIA - TDIA/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA"))//TURNO NOCHE
                                {
                                    TimeSpan Tdia = new TimeSpan(15, 30, 00);//Variable con el horario 19:00 hrs
                                    SalidaTDIA = Convert.ToDateTime(Sabado[0].ToString()).Date + Tdia;//Establece fecha de la checada con la hora de ts
                                                                                                      //----------------------------------------------------------- Sabado ----------------------------------------------------------------------//
                                    if (Sabado.Count == 2)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Sabado[1].ToString()) != null) && (Convert.ToDateTime(Sabado[1].ToString()) >= SalidaTDIA && Convert.ToDateTime(Sabado[1].ToString()) < SalidaTDIA.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[1].ToString()), SalidaTDIA);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Sabado[1].ToString()) != null) && (Convert.ToDateTime(Sabado[1].ToString()) < SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[1].ToString()), SalidaTDIA);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Sabado[1].ToString()) != null) && (Convert.ToDateTime(Sabado[1].ToString()) > SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[1].ToString()), SalidaTDIA);
                                        }
                                    }
                                    if (Sabado.Count >= 4)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Sabado[3].ToString()) != null) && (Convert.ToDateTime(Sabado[3].ToString()) >= SalidaTDIA && Convert.ToDateTime(Sabado[3].ToString()) < SalidaTDIA.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[3].ToString()), SalidaTDIA);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Sabado[3].ToString()) != null) && (Convert.ToDateTime(Sabado[3].ToString()) < SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[3].ToString()), SalidaTDIA);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TDIA") && (Convert.ToDateTime(Sabado[3].ToString()) != null) && (Convert.ToDateTime(Sabado[3].ToString()) > SalidaTDIA))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[3].ToString()), SalidaTDIA);
                                        }
                                    }
                                }//Fin noche
                                 ////////////////////////// - TFIJ - TFIJ - TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ - TFIJ- TFIJ /////////////////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ"))//TURNO NOCHE
                                {
                                    TimeSpan Tfij = new TimeSpan(17, 00, 00);//Variable con el horario 19:00 hrs
                                    SalidaTFIJ = Convert.ToDateTime(Sabado[0].ToString()).Date + Tfij;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Sabado ----------------------------------------------------------------------//
                                    if (Sabado.Count == 2)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Sabado[1].ToString()) != null) && (Convert.ToDateTime(Sabado[1].ToString()) > SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[1].ToString()), SalidaTFIJ);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Sabado[1].ToString()) != null) && (Convert.ToDateTime(Sabado[1].ToString()) < SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[1].ToString()), SalidaTFIJ);
                                        }
                                    }
                                    if (Sabado.Count >= 4)
                                    {
                                        if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Sabado[3].ToString()) != null) && (Convert.ToDateTime(Sabado[3].ToString()) >= SalidaTFIJ && Convert.ToDateTime(Sabado[3].ToString()) < SalidaTFIJ.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                        {
                                            c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[3].ToString()), SalidaTFIJ);//Actualiza la checada al turno asignado
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Sabado[3].ToString()) != null) && (Convert.ToDateTime(Sabado[3].ToString()) < SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[3].ToString()), SalidaTFIJ);
                                        }
                                        else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TFIJ") && (Convert.ToDateTime(Sabado[3].ToString()) != null) && (Convert.ToDateTime(Sabado[3].ToString()) > SalidaTFIJ))
                                        {
                                            c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[3].ToString()), SalidaTFIJ);
                                        }
                                    }
                                }//Fin TFIJ
                                 ///////////////////////// - TNO2 - TNO2 - TNO2 - TNO2 - TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2 - TNO2 - TNO2- TNO2 - TNO2- TNO2 - TNO2- TNO2 ///////////////////////////////////////////////////////////////////////////////
                                if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2"))//TURNO NOCHE
                                {
                                    TimeSpan Tno2 = new TimeSpan(19, 00, 00);//Variable con el horario 19:00 hrs
                                    SalidaTNO2 = Convert.ToDateTime(Sabado[0].ToString()).Date + Tno2;//Establece fecha de la checada con la hora de ts
                                    //----------------------------------------------------------- Sabado ----------------------------------------------------------------------//
                                    if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Sabado[3].ToString()) != null) && (Convert.ToDateTime(Sabado[3].ToString()) >= SalidaTNO2 && Convert.ToDateTime(Sabado[3].ToString()) < SalidaTNO2.AddMinutes(30)))//SALIDA DEL TRABAJO IGUAL O DESPUES DE LA SALIDA, SIENDO MENOR A NOMAS DE 30 MINUTOS
                                    {
                                        c.ActualizarChecada(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[3].ToString()), SalidaTNO2);//Actualiza la checada al turno asignado
                                    }
                                    else if (dt.Rows[indiceTurno]["Turno"].ToString().Equals("TNO2") && (Convert.ToDateTime(Sabado[3].ToString()) != null) && (Convert.ToDateTime(Sabado[3].ToString()) < SalidaTNO2))
                                    {
                                        c.AjusteSalida(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[3].ToString()), SalidaTNO2);
                                    }
                                }//Fin TNO2
                            }
                            catch (Exception ex) {
                                //MessageBox.Show("Sabado Error: " + ex);
                            }//TRY DIA Martes
                            break;
                    }//SWITCH DIA DE LA SEMANA
                }//FOR DE TURNOS
            }//FOR EMPLEADO EN CURSO
            BarraProgreso.Visible = false;
            chekSalida.Checked = true;
            MessageBox.Show("Salidas ajustadas");
        }

        private void Comparador_Click(object sender, EventArgs e)
        {
            DateTime FechaInicio, FechaFinal;
            String FechaFormateada, FinalFormateada;

            FechaInicio = InicioComparador.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro
                                                 // FechaInicio=FechaInicio.AddDays(-1);//Se agrega 1 dia más
            FechaFormateada = FechaInicio.ToString("yyyy-MM-dd 00:00:00:000");

            FechaFinal = FinComparador.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro  
            FechaFinal = FechaFinal.AddDays(1);
            FinalFormateada = FechaFinal.ToString("yyyy-MM-dd 00:00:00:000");

            MessageBox.Show("Fecha inicio: " + FechaInicio + " \n A \n fecha final: " + FechaFinal);
            SqlConnection cn; // para la conexion a SQL
            try
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                //importar la libreria de widws form para abrir una ventana 
                SqlCommand cmd = new SqlCommand("InformeComparador", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FechaInicio", FechaFormateada);
                cmd.Parameters.AddWithValue("@FechaFinal", FinalFormateada);
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(cmd.ExecuteReader());
                gridComparador.DataSource = dt;
            }
            catch (Exception ex)
            {
                //ver.. explorador de servidores... Conectar Base de Datos.. Cambiar... 
                //Microsoft SQL Server... Nombre Servidor (.)...Escoger La Base de Datos
                MessageBox.Show("No se conecto la Base d Datos" + ex.ToString());
            }
        }

        private void butFestivoTodos_Click(object sender, EventArgs e)
        {
            String TipoChecada = listTipoChecada.Text;
            //******************************************* FORMATO DE CHECADA **********************************************************************************/
            String Horas = txtHoraNueva.Text;
            int Hora = Convert.ToInt32(Horas.Substring(0, 2));
            int Minutos = Convert.ToInt32(Horas.Substring(3, 2));
            int Segundos = Convert.ToInt32(0/*Horas.Substring(6,2)*/);
            DateTime NuevaFecha;
            DateTime NuevaChecada;
            String FechaFormateada;
            NuevaFecha = dateChecadaNueva.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro
            FechaFormateada = NuevaFecha.ToString("yyyy-MM-dd 00:00:0");

            TimeSpan TS_Horas = new TimeSpan(Hora, Minutos, Segundos);//Variable con el horario 19:00 hrs
            NuevaChecada = Convert.ToDateTime(Convert.ToDateTime(FechaFormateada) + TS_Horas);//Establece fecha de la checada con la hora de ts
            //**************************************************************************************************************************************************/
            if (listTurno.Text.Equals("NOCHE"))
            {
                int limite = 0;
                while (limite < NOCHE)
                {

                    FC.InsertarFestivos(NuevaChecada, listTipoChecada.Text.ToString(), listTurno.Text.ToString(), "Error", "Festivo");
                    limite++;
                }
            }
            if (listTurno.Text.Equals("TNO2"))
            {
                int limite = 0;
                while (limite < TNO2)
                {
                    FC.InsertarFestivos(NuevaChecada, listTipoChecada.Text.ToString(), listTurno.Text.ToString(), "Error", "Festivo");
                    limite++;
                }
            }
            if (listTurno.Text.Equals("TARDE"))
            {
                int limite = 0;
                while (limite < TARDE)
                {
                    FC.InsertarFestivos(NuevaChecada, listTipoChecada.Text.ToString(), listTurno.Text.ToString(), "Error", "Festivo");
                    limite++;
                }
            }
            if (listTurno.Text.Equals("TDIA"))
            {
                int limite = 0;
                while (limite < TDIA)
                {
                    FC.InsertarFestivos(NuevaChecada, listTipoChecada.Text.ToString(), listTurno.Text.ToString(), "Error", "Festivo");
                    limite++;
                }
            }
            if (listTurno.Text.Equals("TFIJ"))
            {
                int limite = 0;
                while (limite < TFIJ)
                {
                    FC.InsertarFestivos(NuevaChecada, listTipoChecada.Text.ToString(), listTurno.Text.ToString(), "Error", "Festivo");
                    limite++;
                }
            }
        }

        private void butExtras_Click(object sender, EventArgs e)
        {
            if (GridSemana.Rows.Count > 1)
            {
                for (int i = GridSemana.Rows.Count - 1; i >= 0; i--)
                {
                    GridSemana.Rows.RemoveAt(i);
                }
            }
            GridSemana.Refresh();
            DateTime FechaInicio, FechaFinal;
            String FechaFormateada, FinalFormateada;
            FechaInicio = dateInicio.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro
                                           // FechaInicio=FechaInicio.AddDays(-1);//Se agrega 1 dia más
            FechaFormateada = FechaInicio.ToString("yyyy-MM-dd 00:00:00:000");

            FechaFinal = dateFinal.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro  
            FechaFinal = FechaFinal.AddDays(1);
            FinalFormateada = FechaFinal.ToString("yyyy-MM-dd 00:00:00:000");

            DateTime EntradaNoche = Convert.ToDateTime("19:00:00");
            DateTime EntradaSAB = Convert.ToDateTime("08:00:00");
            DateTime EntradaTARDE = Convert.ToDateTime("15:30:00");
            DateTime EntradaTDIA = Convert.ToDateTime("07:00:00");
            DateTime EntradaTFIJ = Convert.ToDateTime("08:00:00");
            DateTime EntradaTNO2 = Convert.ToDateTime("07:00:00");

            ArrayList Lunes = new ArrayList();
            ArrayList Martes = new ArrayList();
            ArrayList Miercoles = new ArrayList();
            ArrayList Jueves = new ArrayList();
            ArrayList Viernes = new ArrayList();
            ArrayList Sabado = new ArrayList();
            ArrayList Semana = new ArrayList { "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado" };
            ArrayList Turnos = new ArrayList { "NOCHE", "SAB", "TARDE", "TDIA", "TFIJ", "TNO2" };
            ArrayList Empleado = new ArrayList();
            ArrayList Tipo = new ArrayList();

            BarraProgreso.Maximum = Turnos.Count;
            BarraProgreso.BackColor = Color.White;
            BarraProgreso.ProgressColor1 = Color.CornflowerBlue;
            BarraProgreso.ProgressColor2 = Color.DarkRed;
            BarraProgreso.Visible = true;
            BarraProgreso.ForeColor = Color.Cyan;
            /////////////////////////////////////GET COLECCIÓN && SELECCIÓN DE EMPLEADOS POR TURNOS /////////////////////////////////////////////////////////////
            for (int t = 0; t < Turnos.Count; t++)
            {
                BarraProgreso.Value = t;
                BarraProgreso.Update();//Contador de progreso
                Empleado.Clear();
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                string query = "SELECT DISTINCT Trabajador FROM Checadas WHERE Turno = '" + Turnos[t].ToString() + "' Order by Trabajador ";

                SqlCommand cmd = new SqlCommand(query, cn);
                DataTable t1 = new DataTable();
                using (SqlDataAdapter a = new SqlDataAdapter(cmd))
                {
                    a.Fill(t1);
                }

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
                    query = "SELECT * from Checadas WHERE Trabajador = '" + Empleado[indiceEmpleados].ToString() + "' AND (TipoChecada = 'SI' OR TipoChecada = 'EI') AND Checada BETWEEN '" + FechaFormateada + "' AND '" + FinalFormateada + "' order by Checada";
                    SqlCommand cmd2 = new SqlCommand(query, cn);
                    DataTable TablaSi = new DataTable();
                    using (SqlDataAdapter a = new SqlDataAdapter(cmd2))
                    {
                        a.Fill(TablaSi);
                    }
                    cn.Close();

                    /////////////////////////////////////////////// REVISAR SALIDA DIARIA //////////////////////////////////////////////////////

                    for (int indiceTurno = 0; indiceTurno < TablaSi.Rows.Count; indiceTurno++)//Recorrer cada registro para obtener el turno
                    {
                        ////////////////////////////////// Codigo para el dia de la semana /////////////////////////////////////////////////////
                        DateTime dateValue = new DateTime();
                        CultureInfo cul = CultureInfo.CurrentCulture;
                        // Usa la fecha formateada y calcula el número de la semana
                        int NumeroSemana = cul.Calendar.GetWeekOfYear(dateValue, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
                        dateValue = Convert.ToDateTime(TablaSi.Rows[indiceTurno]["Checada"].ToString());//Consulto la checada del indice
                        String dia = dateValue.DayOfWeek.ToString();//Convierto la fecha en dia de la semana

                        switch (dia)
                        {
                            case "Lunes":
                            case "Monday":
                                Lunes.Add(TablaSi.Rows[indiceTurno]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                Tipo.Add(TablaSi.Rows[indiceTurno]["TipoChecada"].ToString());
                                break;
                            case "Martes":
                            case "Tuesday":
                                Martes.Add(TablaSi.Rows[indiceTurno]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                break;
                            case "Miércoles":
                            case "Wednesday":
                                Miercoles.Add(TablaSi.Rows[indiceTurno]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                break;
                            case "Jueves":
                            case "Thursday":
                                Jueves.Add(TablaSi.Rows[indiceTurno]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                break;
                            case "Viernes":
                            case "Friday":
                                Viernes.Add(TablaSi.Rows[indiceTurno]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                break;
                            case "Sábado":
                            case "Saturday":
                                Sabado.Add(TablaSi.Rows[indiceTurno]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                break;
                        }//FIN SWITCH dias
                    }//Fin FOR Checada por dias

                    for (int s = 0; s < Semana.Count; s++)//Recorrer cada registro para obtener el turno
                    {
                        switch (Semana[s].ToString())
                        {
                            case "Lunes":
                            case "Monday":
                                if (Lunes.Count >= 1)
                                {
                                    if (Turnos[t].ToString().Equals("NOCHE"))
                                    {
                                        TimeSpan Noche = new TimeSpan(07, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaNoche = Convert.ToDateTime(Lunes[0].ToString()).Date + Noche;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Lunes[Lunes.Count - 1].ToString()) > SalidaNoche)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[Lunes.Count - 1].ToString()), SalidaNoche);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TARDE"))
                                    {
                                        TimeSpan Tarde = new TimeSpan(23, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTarde = Convert.ToDateTime(Lunes[0].ToString()).Date + Tarde;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Lunes[Lunes.Count - 1].ToString()) > SalidaTarde)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[Lunes.Count - 1].ToString()), SalidaTarde);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TDIA"))
                                    {
                                        TimeSpan Tdia = new TimeSpan(15, 30, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTDIA = Convert.ToDateTime(Lunes[0].ToString()).Date + Tdia;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Lunes[Lunes.Count - 1].ToString()) > SalidaTDIA)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[Lunes.Count - 1].ToString()), SalidaTDIA);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TFIJ"))
                                    {
                                        TimeSpan Tfij = new TimeSpan(17, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTFIJ = Convert.ToDateTime(Lunes[0].ToString()).Date + Tfij;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Lunes[Lunes.Count - 1].ToString()) > SalidaTFIJ)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[Lunes.Count - 1].ToString()), SalidaTFIJ);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TNO2"))
                                    {
                                        TimeSpan Tno2 = new TimeSpan(19, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTNO2 = Convert.ToDateTime(Lunes[0].ToString()).Date + Tno2;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Lunes[Lunes.Count - 1].ToString()) > SalidaTNO2)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Lunes[Lunes.Count - 1].ToString()), SalidaTNO2);
                                        }
                                    }
                                }
                                break;
                            case "Martes":
                            case "Tuesday":
                                if (Martes.Count >= 1)
                                {
                                    if (Turnos[t].ToString().Equals("NOCHE"))
                                    {
                                        TimeSpan Noche = new TimeSpan(07, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaNoche = Convert.ToDateTime(Martes[0].ToString()).Date + Noche;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Martes[Martes.Count - 1].ToString()) > SalidaNoche)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[Martes.Count - 1].ToString()), SalidaNoche);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TARDE"))
                                    {
                                        TimeSpan Tarde = new TimeSpan(23, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTarde = Convert.ToDateTime(Martes[0].ToString()).Date + Tarde;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Martes[Martes.Count - 1].ToString()) > SalidaTarde)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[Martes.Count - 1].ToString()), SalidaTarde);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TDIA"))
                                    {
                                        TimeSpan Tdia = new TimeSpan(15, 30, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTDIA = Convert.ToDateTime(Martes[0].ToString()).Date + Tdia;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Martes[Martes.Count - 1].ToString()) > SalidaTDIA)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[Martes.Count - 1].ToString()), SalidaTDIA);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TFIJ"))
                                    {
                                        TimeSpan Tfij = new TimeSpan(17, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTFIJ = Convert.ToDateTime(Martes[0].ToString()).Date + Tfij;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Martes[Martes.Count - 1].ToString()) > SalidaTFIJ)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[Martes.Count - 1].ToString()), SalidaTFIJ);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TNO2"))
                                    {
                                        TimeSpan Tno2 = new TimeSpan(19, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTNO2 = Convert.ToDateTime(Martes[0].ToString()).Date + Tno2;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Martes[Martes.Count - 1].ToString()) > SalidaTNO2)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Martes[Martes.Count - 1].ToString()), SalidaTNO2);
                                        }
                                    }
                                }
                                break;
                            case "Miercoles":
                            case "Wednesday":
                                if (Miercoles.Count >= 1)
                                {
                                    if (Turnos[t].ToString().Equals("NOCHE"))
                                    {
                                        TimeSpan Noche = new TimeSpan(07, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaNoche = Convert.ToDateTime(Miercoles[0].ToString()).Date + Noche;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Miercoles[Miercoles.Count - 1].ToString()) > SalidaNoche)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[Miercoles.Count - 1].ToString()), SalidaNoche);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TARDE"))
                                    {
                                        TimeSpan Tarde = new TimeSpan(23, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTarde = Convert.ToDateTime(Miercoles[0].ToString()).Date + Tarde;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Miercoles[Miercoles.Count - 1].ToString()) > SalidaTarde)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[Miercoles.Count - 1].ToString()), SalidaTarde);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TDIA"))
                                    {
                                        TimeSpan Tdia = new TimeSpan(15, 30, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTDIA = Convert.ToDateTime(Miercoles[0].ToString()).Date + Tdia;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Miercoles[Miercoles.Count - 1].ToString()) > SalidaTDIA)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[Miercoles.Count - 1].ToString()), SalidaTDIA);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TFIJ"))
                                    {
                                        TimeSpan Tfij = new TimeSpan(17, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTFIJ = Convert.ToDateTime(Miercoles[0].ToString()).Date + Tfij;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Miercoles[Miercoles.Count - 1].ToString()) > SalidaTFIJ)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[Miercoles.Count - 1].ToString()), SalidaTFIJ);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TNO2"))
                                    {
                                        TimeSpan Tno2 = new TimeSpan(19, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTNO2 = Convert.ToDateTime(Miercoles[0].ToString()).Date + Tno2;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Miercoles[Miercoles.Count - 1].ToString()) > SalidaTNO2)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Miercoles[Miercoles.Count - 1].ToString()), SalidaTNO2);
                                        }
                                    }
                                }
                                break;
                            case "Jueves":
                            case "Thursday":
                                if (Jueves.Count >= 1)
                                {
                                    if (Turnos[t].ToString().Equals("NOCHE"))
                                    {
                                        TimeSpan Noche = new TimeSpan(07, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaNoche = Convert.ToDateTime(Jueves[0].ToString()).Date + Noche;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Jueves[Jueves.Count - 1].ToString()) > SalidaNoche)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[Jueves.Count - 1].ToString()), SalidaNoche);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TARDE"))
                                    {
                                        TimeSpan Tarde = new TimeSpan(23, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTarde = Convert.ToDateTime(Jueves[0].ToString()).Date + Tarde;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Jueves[Jueves.Count - 1].ToString()) > SalidaTarde)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[Jueves.Count - 1].ToString()), SalidaTarde);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TDIA"))
                                    {
                                        TimeSpan Tdia = new TimeSpan(15, 30, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTDIA = Convert.ToDateTime(Jueves[0].ToString()).Date + Tdia;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Jueves[Jueves.Count - 1].ToString()) > SalidaTDIA)
                                        {
                                             c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[Jueves.Count - 1].ToString()), SalidaTDIA);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TFIJ"))
                                    {
                                        TimeSpan Tfij = new TimeSpan(17, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTFIJ = Convert.ToDateTime(Jueves[0].ToString()).Date + Tfij;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Jueves[Jueves.Count - 1].ToString()) > SalidaTFIJ)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[Jueves.Count - 1].ToString()), SalidaTFIJ);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TNO2"))
                                    {
                                        TimeSpan Tno2 = new TimeSpan(19, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTNO2 = Convert.ToDateTime(Jueves[0].ToString()).Date + Tno2;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Jueves[Jueves.Count - 1].ToString()) > SalidaTNO2)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Jueves[Jueves.Count - 1].ToString()), SalidaTNO2);
                                        }
                                    }
                                }
                                break;
                            case "Viernes":
                            case "Friday":
                                if (Viernes.Count >= 1)
                                {
                                    if (Turnos[t].ToString().Equals("NOCHE"))
                                    {
                                        TimeSpan Noche = new TimeSpan(07, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaNoche = Convert.ToDateTime(Viernes[0].ToString()).Date + Noche;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Viernes[Viernes.Count - 1].ToString()) > SalidaNoche)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[Viernes.Count - 1].ToString()), SalidaNoche);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TARDE"))
                                    {
                                        TimeSpan Tarde = new TimeSpan(23, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTarde = Convert.ToDateTime(Viernes[0].ToString()).Date + Tarde;//Establece fecha de la checada con la hora de ts
                                        if(Convert.ToDateTime(Viernes[Viernes.Count - 1].ToString()) > SalidaTarde)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[Viernes.Count - 1].ToString()), SalidaTarde);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TDIA"))
                                    {
                                        TimeSpan Tdia = new TimeSpan(15, 30, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTDIA = Convert.ToDateTime(Viernes[0].ToString()).Date + Tdia;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Viernes[Viernes.Count - 1].ToString()) > SalidaTDIA)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[Viernes.Count - 1].ToString()), SalidaTDIA);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TFIJ"))
                                    {
                                        TimeSpan Tfij = new TimeSpan(17, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTFIJ = Convert.ToDateTime(Viernes[0].ToString()).Date + Tfij;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Viernes[Viernes.Count - 1].ToString()) > SalidaTFIJ)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[Viernes.Count - 1].ToString()), SalidaTFIJ);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TNO2"))
                                    {
                                        TimeSpan Tno2 = new TimeSpan(19, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTNO2 = Convert.ToDateTime(Viernes[0].ToString()).Date + Tno2;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Viernes[Viernes.Count - 1].ToString()) > SalidaTNO2)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Viernes[Viernes.Count - 1].ToString()), SalidaTNO2);
                                        }
                                    }
                                }
                                break;
                            case "Sabado":
                            case "Saturday":
                                if (Sabado.Count >= 1)
                                {
                                    if (Turnos[t].ToString().Equals("NOCHE"))
                                    {
                                        TimeSpan Noche = new TimeSpan(07, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaNoche = Convert.ToDateTime(Sabado[0].ToString()).Date + Noche;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Sabado[Sabado.Count - 1].ToString()) > SalidaNoche)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[Sabado.Count - 1].ToString()), SalidaNoche);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("SAB"))
                                    {
                                        TimeSpan SAB = new TimeSpan(13, 30, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaSAB = Convert.ToDateTime(Sabado[0].ToString()).Date + SAB;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Sabado[Sabado.Count - 1].ToString()) > SalidaSAB)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[Sabado.Count - 1].ToString()), SalidaSAB);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TARDE"))
                                    {
                                        TimeSpan Tarde = new TimeSpan(23, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTarde = Convert.ToDateTime(Sabado[0].ToString()).Date + Tarde;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Sabado[Sabado.Count - 1].ToString()) > SalidaTarde)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[Sabado.Count - 1].ToString()), SalidaTarde);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TDIA"))
                                    {
                                        TimeSpan Tdia = new TimeSpan(15, 30, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTDIA = Convert.ToDateTime(Sabado[0].ToString()).Date + Tdia;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Sabado[Sabado.Count - 1].ToString()) > SalidaTDIA)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[Sabado.Count - 1].ToString()), SalidaTDIA);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TFIJ"))
                                    {
                                        TimeSpan Tfij = new TimeSpan(17, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTFIJ = Convert.ToDateTime(Sabado[0].ToString()).Date + Tfij;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Sabado[Sabado.Count - 1].ToString()) > SalidaTFIJ)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[Sabado.Count - 1].ToString()), SalidaTFIJ);
                                        }
                                    }
                                    else if (Turnos[t].ToString().Equals("TNO2"))
                                    {
                                        TimeSpan Tno2 = new TimeSpan(19, 00, 00);//Variable con el horario 19:00 hrs
                                        DateTime SalidaTNO2 = Convert.ToDateTime(Sabado[0].ToString()).Date + Tno2;//Establece fecha de la checada con la hora de ts
                                        if (Convert.ToDateTime(Sabado[Sabado.Count - 1].ToString()) > SalidaTNO2)
                                        {
                                            c.AjusteSalidaSI(Empleado[indiceEmpleados].ToString(), Convert.ToDateTime(Sabado[Sabado.Count - 1].ToString()), SalidaTNO2);
                                        }
                                    }
                                }
                                break;
                        }//Fin switch                  
                    }//FOR DE TURNOS
                }//FOR EMPLEADO EN CURSO
            }
            BarraProgreso.Visible = false;
            chekSI.Checked = true;
            MessageBox.Show("Fin de SI");
        }

        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////// CALCULAR SEMANA DE TRABAJADOR /////////////////////////////////////////////////////////
        private void butCalcularSemana_Click(object sender, EventArgs e)
        {
            DateTime SalidaNoche = Convert.ToDateTime("07:00:00");
            DateTime SalidaSAB = Convert.ToDateTime("13:00:00");
            DateTime SalidaTARDE = Convert.ToDateTime("23:00:00");
            DateTime SalidaTDIA = Convert.ToDateTime("15:30:00");
            DateTime SalidaTFIJ = Convert.ToDateTime("17:00:00");
            DateTime SalidaTNO2 = Convert.ToDateTime("19:00:00");

            ArrayList Lunes = new ArrayList();
            ArrayList Martes = new ArrayList();
            ArrayList Miercoles = new ArrayList();
            ArrayList Jueves = new ArrayList();
            ArrayList Viernes = new ArrayList();
            ArrayList Sabado = new ArrayList();
            ArrayList Semana = new ArrayList { "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado" };
            ArrayList Turno = new ArrayList {"NOCHE", "SAB","TARDE","TDIA","TFIJ","TNO2"};
            ArrayList Empleado = new ArrayList();
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            query = "TRUNCATE TABLE Horas";
            cmd = new SqlCommand(query, cn);
            cmd.ExecuteNonQuery();//para que relice la sentencia sql
            cn.Close();

            /////////////////////////////////////GET COLECCIÓN && SELECCIÓN DE EMPLEADOS POR TURNOS ///////////////////////////////
            for (int T = 0; T < 5; T++)
            {
                switch (Turno[T].ToString())
                {
                    case "NOCHE":
                        sm.TurnoNOCHE(this.ListasFiltradas("NOCHE"));
                        break;
                    case "SAB":
                        sm.TurnoSAB(this.ListasFiltradas("SAB"));
                        break;
                    case "TARDE":
                        sm.TurnoTARDE(this.ListasFiltradas("TARDE"));
                        break;
                    case "TDIA":
                        sm.TurnoTDIA(this.ListasFiltradas("TDIA"));
                        break;
                    case "TFIJ":
                        sm.TurnoTFIJ(this.ListasFiltradas("TFIJ"));
                        break;
                    default:
                        break;
                }
            }//FOR DEL EMPLEADO

            cn.Open();
            query = "SELECT DISTINCT Trabajador FROM Checadas Order by Trabajador ";
            cmd = new SqlCommand(query, cn);
            DataTable t1 = new DataTable();
            using (SqlDataAdapter a = new SqlDataAdapter(cmd))
            {
                a.Fill(t1);
            }
            cn.Close();
            for (int x = 0; x < t1.Rows.Count; x++)//FOR[1] : Agrega los Trabajadores a una lista
            {
                cn.Open();
                //importar la libreria de widws form para abrir una ventana 
                cmd = new SqlCommand("GenerarTotal", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Empleado", t1.Rows[x]["Trabajador"].ToString());
                cmd.ExecuteReader();
                cn.Close();
            }
            MessageBox.Show("Fin de cálculo");
        }//Fin de Función
        
        public ArrayList ListasFiltradas(String Turno)
        {
            ArrayList Empleado = new ArrayList();
            cn = new SqlConnection(var.GetConexionDB());
            if (Turno.Equals("NOCHE"))
            {
                Empleado.Clear();
               
                cn.Open();
                string query = "SELECT DISTINCT Trabajador FROM Checadas WHERE Turno = 'NOCHE' Order by Trabajador ";

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
            }else if (Turno.Equals("SAB"))
            {
                Empleado.Clear();
                //cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                string query = "SELECT DISTINCT Trabajador FROM Checadas WHERE Turno = 'SAB' Order by Trabajador ";

                SqlCommand cmd = new SqlCommand(query, cn);
                DataTable t1 = new DataTable();
                using (SqlDataAdapter a = new SqlDataAdapter(cmd))
                {
                    a.Fill(t1);
                }

                for (int x = 0; x < t1.Rows.Count; x++)//FOR[1] : Agrega los Trabajadores a una lista
                {
                    Empleado.Add(t1.Rows[x]["Trabajador"].ToString());
                }
            }
            else if (Turno.Equals("TARDE"))
            {
                Empleado.Clear();
                //cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                string query = "SELECT DISTINCT Trabajador FROM Checadas WHERE Turno = 'TARDE' Order by Trabajador ";

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
            }
            else if (Turno.Equals("TDIA"))
            {
                Empleado.Clear();
                //cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                string query = "SELECT DISTINCT Trabajador FROM Checadas WHERE Turno = 'TDIA' Order by Trabajador ";

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
            }
            else if (Turno.Equals("TFIJ"))
            {
                Empleado.Clear();
                //cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                string query = "SELECT DISTINCT Trabajador FROM Checadas WHERE Turno = 'TFIJ' Order by Trabajador ";

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
            }
            else if (Turno.Equals("TNO2"))
            {
                Empleado.Clear();
                //cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                string query = "SELECT DISTINCT Trabajador FROM Checadas WHERE Turno = 'TNO2' Order by Trabajador ";

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
            }
            return Empleado;
        }//Fin de Listas

        private void ExcelSemana_Click(object sender, EventArgs e, String Empleado)
        {
           DataTable dt = FC.HorasEmpleado();
        }

        private void txtUnificar_Click(object sender, EventArgs e)
        {
            try
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                query = "TRUNCATE TABLE Prenomina";
                SqlCommand cmd = new SqlCommand(query, cn);
                cmd.ExecuteNonQuery();//para que relice la sentencia sql
                cn.Close();

                ArrayList Lunes = new ArrayList();
                ArrayList Martes = new ArrayList();
                ArrayList Miercoles = new ArrayList();
                ArrayList Jueves = new ArrayList();
                ArrayList Viernes = new ArrayList();
                ArrayList Sabado = new ArrayList();
                ArrayList Semana = new ArrayList { "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado" };

                //cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                query = "SELECT distinct C.Trabajador,C.Checada,C.Turno, C.Estatus,C.Evento,C.Estatus, H.Total,H.Lunes,H.Martes,H.Miercoles,H.Jueves,H.Miercoles,H.Jueves,H.Viernes,H.Sabado " +
                        " FROM Checadas C, Horas H WHERE C.Trabajador = H.Empleado";

                cmd = new SqlCommand(query, cn);
                DataTable dt = new DataTable();
                using (SqlDataAdapter a = new SqlDataAdapter(cmd))

                {
                    a.Fill(dt);
                }

                cn.Close();
                Lunes.Clear();
                Martes.Clear();
                Miercoles.Clear();
                Jueves.Clear();
                Viernes.Clear();
                Viernes.Clear();
                Sabado.Clear();
                /////////////////////////////////////////////// OBTENER CHECADAS DEL DIA //////////////////////////////////////////////////////
                for (int reg = 0; reg < dt.Rows.Count; reg++)
                {
                    DateTime dateValue = new DateTime();
                    CultureInfo cul = CultureInfo.CurrentCulture;
                    String query = "";
                    String Trabajador = dt.Rows[reg]["Trabajador"].ToString();
                    String Checada = dt.Rows[reg]["Checada"].ToString();
                    String Turno = dt.Rows[reg]["Turno"].ToString();
                    String Estatus = dt.Rows[reg]["Estatus"].ToString();
                    String Evento = dt.Rows[reg]["Evento"].ToString();
                    Double Total = Convert.ToDouble(dt.Rows[reg]["Total"].ToString());
                    String Lun = dt.Rows[reg]["Lunes"].ToString();
                    String Mar = dt.Rows[reg]["Martes"].ToString();
                    String Mier = dt.Rows[reg]["Miercoles"].ToString();
                    String Jue = dt.Rows[reg]["Jueves"].ToString();
                    String Vie = dt.Rows[reg]["Viernes"].ToString();
                    String Sa = dt.Rows[reg]["Sabado"].ToString();

                    // Usa la fecha formateada y calcula el número de la semana
                    int NumeroSemana = cul.Calendar.GetWeekOfYear(dateValue, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
                    dateValue = Convert.ToDateTime(dt.Rows[reg]["Checada"].ToString());//Consulto la checada del indice
                    String dia = dateValue.DayOfWeek.ToString();//Convierto la fecha en dia de la semana
                    switch (dia)
                    {
                        case "Lunes":
                        case "Monday":
                            cn = new SqlConnection(var.GetConexionDB());
                            cn.Open();
                            query = "INSERT Prenomina (Trabajador,Turno,Estatus,Evento,Total,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,TLunes)VALUES('" + Trabajador + "','" + Turno + "','" + Estatus + "','" + Evento + "'," + Total + ",'" + Lun + "','" + Mar + "','" + Mier + "','" + Jue + "','" + Vie + "','" + Sa + "','" + Turno +"')";
                            cmd = new SqlCommand(query, cn);
                            cmd.ExecuteNonQuery();//para que relice la sentencia sql
                            cn.Close();
                            break;
                        case "Martes":
                        case "Tuesday":
                            cn = new SqlConnection(var.GetConexionDB());
                            cn.Open();
                            query = "INSERT Prenomina (Trabajador,Turno,Estatus,Evento,Total,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,TMartes)VALUES('" + Trabajador + "','" + Turno + "','" + Estatus + "','" + Evento + "'," + Total + ",'" + Lun + "','" + Mar + "','" + Mier + "','" + Jue + "','" + Vie + "','" + Sa + "','" + Turno + "')";
                            cmd = new SqlCommand(query, cn);
                            cmd.ExecuteNonQuery();//para que relice la sentencia sql
                            cn.Close();
                            break;
                        case "Miércoles":
                        case "Wednesday":
                            cn = new SqlConnection(var.GetConexionDB());
                            cn.Open();
                            query = "INSERT Prenomina (Trabajador,Turno,Estatus,Evento,Total,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,TMiercoles)VALUES('" + Trabajador + "','" + Turno + "','" + Estatus + "','" + Evento + "'," + Total + ",'" + Lun + "','" + Mar + "','" + Mier + "','" + Jue + "','" + Vie + "','" + Sa + "','" + Turno + "')";
                            cmd = new SqlCommand(query, cn);
                            cmd.ExecuteNonQuery();//para que relice la sentencia sql
                            cn.Close();
                            break;
                        case "Jueves":
                        case "Thursday":
                            cn = new SqlConnection(var.GetConexionDB());
                            cn.Open();
                            query = "INSERT Prenomina (Trabajador,Turno,Estatus,Evento,Total,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,TJueves)VALUES('" + Trabajador + "','" + Turno + "','" + Estatus + "','" + Evento + "'," + Total + ",'" + Lun + "','" + Mar + "','" + Mier + "','" + Jue + "','" + Vie + "','" + Sa + "','" + Turno + "')";
                            cmd = new SqlCommand(query, cn);
                            cmd.ExecuteNonQuery();//para que relice la sentencia sql
                            cn.Close();
                            break;
                        case "Viernes":
                        case "Friday":
                            cn = new SqlConnection(var.GetConexionDB());
                            cn.Open();
                            query = "INSERT Prenomina (Trabajador,Turno,Estatus,Evento,Total,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,TViernes)VALUES('" + Trabajador + "','" + Turno + "','" + Estatus + "','" + Evento + "'," + Total + ",'" + Lun + "','" + Mar + "','" + Mier + "','" + Jue + "','" + Vie + "','" + Sa + "','" + Turno + "')";
                            cmd = new SqlCommand(query, cn);
                            cmd.ExecuteNonQuery();//para que relice la sentencia sql
                            cn.Close();
                            break;
                        case "Sábado":
                        case "Saturday":
                            cn = new SqlConnection(var.GetConexionDB());
                            cn.Open();
                            query = "INSERT Prenomina (Trabajador,Turno,Estatus,Evento,Total,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,TSabado)VALUES('" + Trabajador + "','" + Turno + "','" + Estatus + "','" + Evento + "'," + Total + ",'" + Lun + "','" + Mar + "','" + Mier + "','" + Jue + "','" + Vie + "','" + Sa + "','" + Turno + "')";
                            cmd = new SqlCommand(query, cn);
                            cmd.ExecuteNonQuery();//para que relice la sentencia sql
                            cn.Close();
                            break;
                    }//FIN SWITCH dias
                }//Fin FOR Checada por dias
                MessageBox.Show("Tabla Checadas y Tabla Horas Trabajadas Unidas");
            }catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ExcelSemana_Click(object sender, EventArgs e)
        {
            try
            {
                try
                {
                    cn = new SqlConnection(var.GetConexionDB());
                    cn.Open();
                    SqlCommand cmdo = new SqlCommand("Truncate Table Extras", cn);
                    cmdo.ExecuteNonQuery();//para que relice la sentencia sql
                    cn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("No se conecto" + ex.ToString());
                }

                double TotalSemana = 0;
                double HorasExtras = 0;
                ArrayList Empleado = new ArrayList();

                Empleado.Clear();
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                string query = "SELECT DISTINCT Trabajador FROM Prenomina Order by Trabajador";

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

                for (int emp = 0; emp < Empleado.Count; emp++)
                {
                    HorasExtras = 0;
                    Double HorasTrabajadas = 0;
                    TotalSemana = 0;
                    cn = new SqlConnection(var.GetConexionDB());
                    cn.Open();
                    query = "SELECT DISTINCT * From Prenomina WHERE Trabajador = '" + Empleado[emp].ToString() + "'";

                    cmd = new SqlCommand(query, cn);
                    DataTable dt = new DataTable();
                    
                    using (SqlDataAdapter a = new SqlDataAdapter(cmd))
                    {
                        a.Fill(dt);
                    }
                    cn.Close();

                    for (int x = 0; x < dt.Rows.Count; x++)
                    {
                        String Trabajador = dt.Rows[x]["Trabajador"].ToString();
                        //String Checada = dt.Rows[x]["Checada"].ToString();
                        String Turno = dt.Rows[x]["Turno"].ToString();
                        String Estatus = dt.Rows[x]["Estatus"].ToString();
                        String Evento = dt.Rows[x]["Evento"].ToString();
                        Double Total = Convert.ToDouble(dt.Rows[x]["Total"].ToString());
                        Double Lun = Convert.ToDouble(dt.Rows[x]["Lunes"].ToString());
                        Double Mar = Convert.ToDouble(dt.Rows[x]["Martes"].ToString());
                        Double Mier = Convert.ToDouble(dt.Rows[x]["Miercoles"].ToString());
                        Double Jue = Convert.ToDouble(dt.Rows[x]["Jueves"].ToString());
                        Double Vie = Convert.ToDouble(dt.Rows[x]["Viernes"].ToString());
                        Double Sa = Convert.ToDouble(dt.Rows[x]["Sabado"].ToString());
                        String Lunes = dt.Rows[x]["TLunes"].ToString();
                        String Martes = dt.Rows[x]["TMartes"].ToString();
                        String Miercoles = dt.Rows[x]["TMiercoles"].ToString();
                        String Jueves = dt.Rows[x]["TJueves"].ToString();
                        String Viernes = dt.Rows[x]["TViernes"].ToString();
                        String Sabado = dt.Rows[x]["TSabado"].ToString();
                        String DiaConTurno = "";
                        String VarianteTFIJ = "Entre";
                       
                        if (!Lunes.Equals("/") || (!Lunes.Equals("/") && Evento.Equals("Falta")))//Si asi se deberia 
                        {
                            DiaConTurno = Lunes;
                            HorasTrabajadas = HorasTrabajadas + Lun;
                            if (Trabajador.Equals("444"))
                            {
                                MessageBox.Show("Horas Trabajadas Lunes: " + Lun.ToString() + " Turno: " + DiaConTurno);
                            }
                        }
                        else if (!Martes.Equals("/") || (!Martes.Equals("/") && Evento.Equals("Falta")))
                        {
                            DiaConTurno = Martes;
                            HorasTrabajadas = HorasTrabajadas + Mar;
                            if (Trabajador.Equals("444"))
                            {
                                MessageBox.Show("Horas Trabajadas Martes: " + Mar.ToString() + " Turno: " + DiaConTurno);
                            }
                        }
                        else if (!Miercoles.Equals("/") || (!Miercoles.Equals("/") && Evento.Equals("Falta")))
                        {
                            DiaConTurno = Miercoles;
                            HorasTrabajadas = HorasTrabajadas + Mier;
                            if (Trabajador.Equals("444"))
                            {
                                MessageBox.Show("Horas Trabajadas Miercoles: " + Mier.ToString() + " Turno: " + DiaConTurno);
                            }
                        }
                        else if (!Jueves.Equals("/") || (!Jueves.Equals("/") && Evento.Equals("Falta")))
                        {
                            DiaConTurno = Jueves;
                            HorasTrabajadas = HorasTrabajadas + Jue;
                            if (Trabajador.Equals("444"))
                            {
                                MessageBox.Show("Horas Trabajadas Jueves: " + Jue.ToString() + " Turno: " + DiaConTurno);
                            }
                        }
                        else if (!Viernes.Equals("/") || (!Viernes.Equals("/") && Evento.Equals("Falta")))
                        {
                            DiaConTurno = Viernes;
                            HorasTrabajadas = HorasTrabajadas + Vie;
                            if (Trabajador.Equals("444"))
                            {
                                MessageBox.Show("Horas Trabajadas Viernes: " + Vie.ToString() + " Turno: " + DiaConTurno);
                            }
                        }
                        else if (!Sabado.Equals("/") || (!Sabado.Equals("/") && Evento.Equals("Falta")))
                        {
                            DiaConTurno = Sabado;
                            HorasTrabajadas = HorasTrabajadas + Sa;
                            if (DiaConTurno.Equals("TFIJ"))
                            {
                                VarianteTFIJ = "Sabado";
                            }
                            if (Trabajador.Equals("444"))
                            {
                                MessageBox.Show("Horas Trabajadas Sabado: " + Sa.ToString() + " Turno: " + DiaConTurno);
                            }
                        }
                        switch (DiaConTurno)
                        {
                            case "NOCHE":
                                TotalSemana = TotalSemana + 7;
                                break;
                            case "TARDE":
                                TotalSemana = TotalSemana + 7.5;
                                break;
                            case "TDIA":
                                TotalSemana = TotalSemana + 8;
                                if (Trabajador.Equals("444"))
                                {
                                    MessageBox.Show("Horas Trabajadas Semana: " + TotalSemana.ToString());
                                }
                                break;
                            case "TFIJ":
                                if (VarianteTFIJ.Equals("Entre"))
                                {
                                    TotalSemana = TotalSemana + 8.5;
                                }
                                else
                                {
                                    TotalSemana = TotalSemana + 5.5;
                                }
                                break;
                            case "SAB":
                                TotalSemana = TotalSemana + 5.5;
                                break;
                        }
                    }//For DataTable
                   
                    HorasExtras = HorasTrabajadas - TotalSemana;
                    if (HorasExtras > 18)
                    {
                        HorasExtras = 18;
                    }
                    
                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    cn = new SqlConnection(var.GetConexionDB());
                    cn.Open();
                    query = "INSERT Extras (Empleado,Horas)VALUES('" + Empleado[emp].ToString() + "'," + HorasExtras + ")";
                    cmd = new SqlCommand(query, cn);
                    cmd.ExecuteNonQuery();//para que relice la sentencia sql
                    cn.Close();
                }//Termina For Con todas las horas de la semana ya enviando el trabajador y horas al excel
                this.EditarExcel();
            }catch(Exception ex) { MessageBox.Show("No se genero excel" + ex.ToString()); }
        }

        private void butVerSemanaEmpleadoHorario_Click(object sender, EventArgs e)
        {
            if (txtEmpleadoHorario.Text.Equals("#Empleado"))
            {
                MessageBox.Show("No hay empleado a buscar");
            }
            else
            {
                if (GridSemanaHorarios.Columns.Count != 0)//Limpia El Grid
                {
                    GridSemanaHorarios.Rows.Clear();
                    GridSemanaHorarios.Refresh();
                    int i = 0;
                    while (i <= GridSemanaHorarios.Columns.Count)
                    {
                        GridSemanaHorarios.Columns.RemoveAt(0);
                        i = +1;
                    }
                }
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////// CONSULTA SEMANA DE LOS TRABAJADORES /////////////////////////////////////////////////////////
                DateTime FechaInicio, FechaFinal;
                String FechaFormateada, FinalFormateada;

                FechaInicio = TimesHorarioInicio.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro
                                                     // FechaInicio=FechaInicio.AddDays(-1);//Se agrega 1 dia más
                FechaFormateada = FechaInicio.ToString("yyyy-MM-dd 00:00:00:000");

                FechaFinal = TimesHorarioFinal.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro  
                FechaFinal = FechaFinal.AddDays(1);
                FinalFormateada = FechaFinal.ToString("yyyy-MM-dd 00:00:00:000");

                MessageBox.Show("Fecha inicio: " + FechaInicio + " \n A \n fecha final: " + FechaFinal);
                try
                {
                    cn = new SqlConnection(var.GetConexionDB());
                    cn.Open();
                    //importar la libreria de widws form para abrir una ventana 
                    SqlCommand cmd = new SqlCommand("SemanaEmpleado", cn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@FechaInicio", FechaFormateada);
                    cmd.Parameters.AddWithValue("@FechaFinal", FinalFormateada);
                    cmd.Parameters.AddWithValue("@Empleado", txtEmpleadoHorario.Text.ToString());
                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.Load(cmd.ExecuteReader());

                    FC.ObetenrDias(dt);//Obtiene los checadas de cada dia
                    var resultado = FC.ObetenrDias(dt);//Contiene la coleccion de ArrayList de la semana
                    ArrayList Lunes = resultado.Item1;
                    ArrayList Martes = resultado.Item2;
                    ArrayList Miercoles = resultado.Item3;
                    ArrayList Jueves = resultado.Item4;
                    ArrayList Viernes = resultado.Item5;
                    ArrayList Sabado = resultado.Item6;

                    this.Columnas(GridSemanaHorarios);

                    GridSemanaHorarios.Rows.Add();
                    for (int i = 0; i < 6; i++)
                    {
                        GridSemanaHorarios.Rows[i].Cells["Trabajador"].Value = dt.Rows[i]["Trabajador"].ToString();
                       // GridSemanaHorarios.Rows[i].Cells["Turno"].Value = dt.Rows[i]["Turno"].ToString();
                        GridSemanaHorarios.Rows[i].Cells["Estatus"].Value = dt.Rows[i]["Estatus"].ToString();
                        GridSemanaHorarios.Rows.Add();
                    }
                    for (int elementos = 0; elementos < Lunes.Count; elementos++)
                    {
                        if (Lunes[elementos].ToString() != null)
                        {
                            GridSemanaHorarios.Rows[elementos].Cells["Lunes"].Value = Lunes[elementos].ToString();//Se crea fila de [Empleado y tipo checada]
                        }
                    }
                    for (int elementos = 0; elementos < Martes.Count; elementos++)
                    {
                        if (Martes[elementos].ToString() != null)
                        {
                            GridSemanaHorarios.Rows[elementos].Cells["Martes"].Value = Martes[elementos].ToString();//Se crea fila de [Empleado y tipo checada]
                        }
                    }
                    for (int elementos = 0; elementos < Miercoles.Count; elementos++)
                    {
                        if (Miercoles[elementos].ToString() != null)
                        {
                            GridSemanaHorarios.Rows[elementos].Cells["Miércoles"].Value = Miercoles[elementos].ToString();//Se crea fila de [Empleado y tipo checada]
                        }
                    }
                    for (int elementos = 0; elementos < Jueves.Count; elementos++)
                    {
                        if (Jueves[elementos].ToString() != null)
                        {
                            GridSemanaHorarios.Rows[elementos].Cells["Jueves"].Value = Jueves[elementos].ToString();//Se crea fila de [Empleado y tipo checada]
                        }
                    }
                    for (int elementos = 0; elementos < Viernes.Count; elementos++)
                    {
                        if (Viernes[elementos].ToString() != null)
                        {
                            GridSemanaHorarios.Rows[elementos].Cells["Viernes"].Value = Viernes[elementos].ToString();//Se crea fila de [Empleado y tipo checada]
                        }
                    }
                    for (int elementos = 0; elementos < Sabado.Count; elementos++)
                    {
                        if (Sabado[elementos].ToString() != null)
                        {
                            GridSemanaHorarios.Rows[elementos].Cells["Sábado"].Value = Sabado[elementos].ToString();//Se crea fila de [Empleado y tipo checada]
                        }
                    }

                    TLunes.SelectedItem = dt.Rows[0]["Turno"].ToString();
                    TMartes.SelectedItem = dt.Rows[0]["Turno"].ToString();
                    Tmiercoles.SelectedItem = dt.Rows[0]["Turno"].ToString();
                    TJueves.SelectedItem = dt.Rows[0]["Turno"].ToString();
                    TViernes.SelectedItem = dt.Rows[0]["Turno"].ToString();
                    TSabado.SelectedItem = dt.Rows[0]["Turno"].ToString();
                }
                catch (Exception ex) { MessageBox.Show("Error" + ex); }
            }
        }

        private void butActualizarHorario_Click(object sender, EventArgs e)
        {
            try
            {
                String Lunes = TLunes.Text.ToString();
                String Martes = TMartes.Text.ToString();
                String Miercoles = Tmiercoles.Text.ToString();
                String Jueves = TJueves.Text.ToString();
                String Viernes = TViernes.Text.ToString();
                String Sabado = TSabado.Text.ToString();

                String TurnoSemana = TSemana.Text.ToString();
                if (rdbSemanaCompleta.Checked == true)
                {
                    cn = new SqlConnection(var.GetConexionDB());
                    cn.Open();
                    SqlCommand cmd = new SqlCommand("UPDATE Checadas SET Turno = '" + TurnoSemana + "' WHERE Trabajador = '" + txtEmpleadoHorario.Text.ToString() + "'", cn);
                    cmd.ExecuteNonQuery();//para que relice la sentencia sql
                    MessageBox.Show("Turno actualizado");
                    cn.Close();
                }
                else
                {
                    cn = new SqlConnection(var.GetConexionDB());
                    cn.Open();
                    cmd = new SqlCommand("CambioDeHorario", cn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    //cmd.Parameters.AddWithValue("@Checada", Checada);
                    cmd.Parameters.AddWithValue("@Lunes", Lunes);
                    cmd.Parameters.AddWithValue("@Martes", Martes);
                    cmd.Parameters.AddWithValue("@Miercoles", Miercoles);
                    cmd.Parameters.AddWithValue("@Jueves", Jueves);
                    cmd.Parameters.AddWithValue("@Viernes", Viernes);
                    cmd.Parameters.AddWithValue("@Sabado", Sabado);
                    cmd.Parameters.AddWithValue("@Empleado", txtEmpleadoHorario.Text.ToString());
                    //cmd.Parameters.AddWithValue("@Fecha", Fecha);
                    cmd.ExecuteReader();
                    cn.Close();
                    MessageBox.Show("Turno actualizado");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error falta: " + ex.ToString());
            }
        }

        private void txtEmpleadoHorario_Enter(object sender, EventArgs e)
        {
            txtEmpleadoHorario.Text = "";
        }

        private void txtEmpleadoHorario_Leave(object sender, EventArgs e)
        {
            if (txtEmpleadoHorario.Text.Equals(""))
            {
                txtEmpleadoHorario.Text = "#Empleado";
            }
        }

        private void ButCambiarTurno_Click(object sender, EventArgs e)
        {
            DateTime FechaInicio, FechaFinal;
            String FechaFormateada, FinalFormateada;

            FechaInicio = dateChecadaNueva.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro
                                                // FechaInicio=FechaInicio.AddDays(-1);//Se agrega 1 dia más
            FechaFormateada = FechaInicio.ToString("yyyy-MM-dd 00:00:00:000");

            FechaFinal = dateChecadaNueva.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro  
            FechaFinal = FechaFinal.AddDays(1);
            FinalFormateada = FechaFinal.ToString("yyyy-MM-dd 00:00:00:000");

            FC.CambiarTurno(listTurno.Text.ToString(),txtEmpleado.Text.ToString(),FechaFormateada,FinalFormateada);
        }

        public void EditarExcel()
        {
            try
            {
                DateTime FechaInicio, FechaFinal;
                String FechaFormateada, FinalFormateada,FechaPago;

                FechaInicio = dateInicio.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro            
                FechaFormateada = FechaInicio.ToString("dd/MM/yyyy");

                FechaFinal = dateFinal.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro  
                FechaFinal = FechaFinal.AddDays(1);
                FinalFormateada = FechaFinal.ToString("dd/MM/yyyy");

                FechaPago = FechaFinal.AddDays(5).ToShortDateString();

                GregorianCalendar cal = new GregorianCalendar(GregorianCalendarTypes.Localized);
                int NumeroSemana = Convert.ToInt32(cal.GetWeekOfYear(FechaInicio, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday).ToString()) + 1;

                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                query = "SELECT * FROM Extras";

                cmd = new SqlCommand(query, cn);
                DataTable dt = new DataTable();
                using (SqlDataAdapter a = new SqlDataAdapter(cmd))
                {
                    a.Fill(dt);
                }
                cn.Close();
                //////////////////////////////////////////////////////////// INICIA EXCEL /////////////////////////////////////////////////////////////////////////////////////////
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;//Libro de Excel
                Excel.Worksheet xlWorkSheet;//Hoja de excel
                Excel.Range range;//Rango

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(@"c:/E2.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);//se obtiene la hoja 1

                range = xlWorkSheet.UsedRange;
                int rw = range.Rows.Count;//Cantidad FIlas
                int cl = range.Columns.Count;//Cantidad Columnas
                int clExtras = 15;//Columna en la que inicia a insertar horas Extras

                // FC.HorasEmpleado();
                for (int emp = 0; emp < dt.Rows.Count; emp ++)
                {
                    for (int rCnt = 10; rCnt <= rw; rCnt++)//FOR para Filas
                    {
                        String str = Convert.ToString((range.Cells[rCnt, 1] as Excel.Range).Value2);
                        if (str != null)
                        {
                            if (str.Equals(dt.Rows[emp]["Empleado"].ToString()))
                            {
                                double Horas = Convert.ToDouble(dt.Rows[emp]["Horas"].ToString());//Se obtiene el valor de horas Extras
                                while (Horas>0)
                                {
                                    if(Horas >= 3)
                                    {
                                        xlWorkSheet.Cells[rCnt, clExtras] = 3;//INSERTAR INFORMACION EN CELDAS
                                        Horas = Horas - 3;
                                        clExtras++;
                                    }
                                    else if(Horas > 0 && Horas < 3)
                                    {
                                        xlWorkSheet.Cells[rCnt, clExtras] = Horas;//INSERTAR INFORMACION EN CELDAS
                                        Horas = 0;
                                        clExtras++;
                                    }
                                }//Fin de while
                                clExtras = 15;
                                //xlWorkSheet.Cells[rCnt, 15] = dt.Rows[emp]["Horas"].ToString();//INSERTAR INFORMACION EN CELDAS
                            }//IF COMPARADOR 
                        }//IF DE NULLOS
                    }//LISTA DE EMPLEADOS EN EXCEL
                }//FOR DE EMPLEADOS

                xlWorkSheet.Cells[6, 1] = "Periodo No. " + NumeroSemana + " del " + FechaFormateada + " al " +  FinalFormateada  + " Pago: " + FechaPago;//INSERTAR INFORMACION EN CELDAS


                String ruta = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);// Obtiene path de escritorio
                xlWorkBook.SaveAs(ruta + "\\HorasExtras");
                xlWorkBook.Close(true, null, null);
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

                Marshal.ReleaseComObject(xlWorkSheet);//NI IDEA
                Marshal.ReleaseComObject(xlWorkBook);//NI IDEA
                Marshal.ReleaseComObject(xlApp);//NI IDEA
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex);
            }
        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //////////////////////////////////////////////////////////////////// C H E C A D A S - B A Ñ O S //////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 
        private void ButCargarBaños_Click(object sender, EventArgs e)
        {
            try
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                cmd = new SqlCommand("truncate table baños", cn);
                cmd.ExecuteNonQuery();//para que relice la sentencia sql
                cn.Close();

                OpenFileDialog dialog = new OpenFileDialog();

                dialog.Filter = "Archivos de Excel (*.xls;*.xlsx)|*.xls;*.xlsx"; //le indicamos el tipo de filtro en este caso que busque solo los archivos excel
                dialog.Title = "Seleccione el archivo de Excel";//le damos un titulo a la ventana
                dialog.FileName = String.Empty;//inicializamos con vacio el nombre del archivo
                                               //si al seleccionar el archivo damos Ok
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    //el nombre del archivo sera asignado al textbox
                    String Ruta = dialog.FileName;
                    String hoja = "data"; //la variable hoja tendra el valor del textbox donde colocamos el nombre de la hoja
                    ImportarExcel(Ruta, hoja); //se manda a llamar al metodo
                }
                MessageBox.Show("Cargados");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex);
            }
        }

        private void ButBañosOriginal_Click(object sender, EventArgs e)
        {
            gridBaños.DataSource = OPB.ConsultarBaños();
        }

        private void ButCalcularBaños_Click(object sender, EventArgs e)
        {
            OPB.DuplicadosBaños();
            OPB.CalcularBaños();
            gridBaños.DataSource = OPB.ConsultarHorasBaños();

            gridBaños.Columns["id"].Visible = false;  //Esconde Estatus
            gridBaños.Columns["Dia"].Visible = false;  //Esconde Estatus
        }

        private void ButDuplicadosBaños_Click(object sender, EventArgs e)
        {
            OPB.DuplicadosBaños();
        }

        private void PictureBox2_Click(object sender, EventArgs e)
        {
            OPB.ExportarExcelLineas(gridBaños);
        }

        private void AplicarFalta(int DiasSumados,String Evento)
        {
            String TipoChecada = listTipoChecada.Text;
            //******************************************* FORMATO DE CHECADA **********************************************************************************/
            String Horas = "12:00:00";
            int Hora = Convert.ToInt32(Horas.Substring(0, 2));
            int Minutos = Convert.ToInt32(Horas.Substring(3, 2));
            int Segundos = Convert.ToInt32(0/*Horas.Substring(6,2)*/);
            String FechaFormateada;
            DateTime FechaSeleccionada = Convert.ToDateTime(CalendarioIncidencias.SelectionStart.AddDays(DiasSumados).ToShortDateString().ToString());
            FechaFormateada = FechaSeleccionada.ToString("yyyy-MM-dd 00:00:0");

            TimeSpan TS_Horas = new TimeSpan(Hora, Minutos, Segundos);//Variable con el horario 19:00 hrs
            FechaSeleccionada = Convert.ToDateTime(Convert.ToDateTime(FechaFormateada) + TS_Horas);//Establece fecha de la checada con la hora de ts
            //**************************************************************************************************************************************************/
            if (ListaTurnosIncidencias.Text.Equals("NOCHE"))
            {
                int limite = 0;
                while (limite < NOCHE)
                {
                    FC.InsertarChecada(FechaSeleccionada, txtEmpleadoIncidencia.Text, "ET", ListaTurnosIncidencias.Text.ToString(), "Error", Evento);
                    limite++;
                }
            }
            if (ListaTurnosIncidencias.Text.Equals("TNO2"))
            {
                int limite = 0;
                while (limite < TNO2)
                {
                    FC.InsertarChecada(FechaSeleccionada, txtEmpleadoIncidencia.Text, "ET", ListaTurnosIncidencias.Text.ToString(), "Error", Evento);
                    limite++;
                }
            }
            if (ListaTurnosIncidencias.Text.Equals("TARDE"))
            {
                int limite = 0;
                while (limite < TARDE)
                {
                    FC.InsertarChecada(FechaSeleccionada, txtEmpleadoIncidencia.Text, "ET", ListaTurnosIncidencias.Text.ToString(), "Error", Evento);
                    limite++;
                }
            }
            if (ListaTurnosIncidencias.Text.Equals("TDIA"))
            {
                int limite = 0;
                while (limite < TDIA)
                {
                    FC.InsertarChecada(FechaSeleccionada, txtEmpleadoIncidencia.Text, "ET", ListaTurnosIncidencias.Text.ToString(), "Error", Evento);
                    limite++;
                }
            }
            if (ListaTurnosIncidencias.Text.Equals("TFIJ"))
            {
                int limite = 0;
                while (limite < TFIJ)
                {
                    FC.InsertarChecada(FechaSeleccionada, txtEmpleadoIncidencia.Text, "ET", ListaTurnosIncidencias.Text.ToString(), "Error", Evento);
                    limite++;
                }
            }
        }

        private void ButRegistrarFalta_Click(object sender, EventArgs e)
        {
            DateTime FechaInicio = Convert.ToDateTime(CalendarioIncidencias.SelectionStart.ToShortDateString().ToString());//Se crea parametro para lanzarlo a la funcion insertarRegistro
            DialogResult result = MessageBox.Show("Asignar falta: " + FechaInicio.ToShortDateString(), "Salir", MessageBoxButtons.YesNoCancel);
            switch (result)
            {
                case DialogResult.Yes:
                    this.ProcesoFaltaVacacion("Falta");
                    MessageBox.Show("Faltas aplicadas");
                    break;
                case DialogResult.No:
                    break;
                case DialogResult.Cancel:
                    break;
            }
        }

        private void ProcesoFaltaVacacion(String Evento)
        {
            if (ListaTurnosIncidencias.Text.Equals(""))
            {
                MessageBox.Show("Seleeciona Turno");
            }
            else
            {
                try
                {
                    DateTime FechaInicio, FechaFinal;

                    FechaInicio = Convert.ToDateTime(CalendarioIncidencias.SelectionStart.ToShortDateString().ToString());//Se crea parametro para lanzarlo a la funcion insertarRegistro
                    FechaFinal = Convert.ToDateTime(CalendarioIncidencias.SelectionEnd.AddDays(1).ToShortDateString().ToString());//Se crea parametro para lanzarlo a la funcion insertarRegistro

                    // Difference in days, hours, and minutes.
                    TimeSpan ts = FechaFinal - FechaInicio;

                    // Difference in days.
                    int DiferenciaDias = ts.Days;
                    for (int x = 0; x < DiferenciaDias; x++)
                    {
                        this.AplicarFalta(x, Evento);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        private void ButVacacionesIncidencias_Click(object sender, EventArgs e)
        {
            DateTime FechaInicio = Convert.ToDateTime(CalendarioIncidencias.SelectionStart.ToShortDateString().ToString());//Se crea parametro para lanzarlo a la funcion insertarRegistro
            DialogResult result = MessageBox.Show("Asignar vacaciones: " + FechaInicio.ToShortDateString(), "Salir", MessageBoxButtons.YesNoCancel);
            switch (result)
            {
                case DialogResult.Yes:
                    this.ProcesoFaltaVacacion("Vacaciones");
                    MessageBox.Show("Vacaciones registradas");
                    break;
                case DialogResult.No:
                    break;
                case DialogResult.Cancel:
                    break;
            }
        }

        private void ProcesoAsuetoR()
        {
            ArrayList Empleados = new ArrayList();
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            query = "SELECT DISTINCT Trabajador FROM Checadas ORDER BY Trabajador";
            SqlCommand cmd2 = new SqlCommand(query, cn);
            DataTable t1 = new DataTable();
            using (SqlDataAdapter a = new SqlDataAdapter(cmd2))
            {
                a.Fill(t1);
            }

            DateTime FechaInicio = Convert.ToDateTime(CalendarioIncidencias.SelectionStart.ToShortDateString().ToString());//Se crea parametro para lanzarlo a la funcion insertarRegistro

            for (int x = 0; x < t1.Rows.Count; x++)
            {
                Empleados.Add(t1.Rows[x]["Trabajador"].ToString());
            }
            cn.Close();

            for (int i = 0; i < Empleados.Count; i++)
            {
                this.InsertarIncidencia(FechaInicio, Empleados[i].ToString());
            }
        }

        private void InsertarIncidencia(DateTime FechaAsueto, String Empleados)//Inserta las vacaciones
        {
            //******************************************* TURNO DE EMPLEADO **********************************************************************************/
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            query = "SELECT DISTINCT Turno,Checada FROM Checadas WHERE Trabajador = '" + Empleados + "' ORDER BY Checada";
            SqlCommand cmd2 = new SqlCommand(query, cn);
            DataTable t1 = new DataTable();
            using (SqlDataAdapter a = new SqlDataAdapter(cmd2))
            {
                a.Fill(t1);
            }
            String Turno = t1.Rows[0]["Turno"].ToString();
            cn.Close();
            //******************************************* FORMATO DE CHECADA **********************************************************************************/
            DateTime EntradaT = new DateTime();
            DateTime SalidaT = new DateTime();
            DateTime EntradaC = new DateTime();
            DateTime SalidaC = new DateTime();

            DateTime NuevaChecada;
            String FechaFormateada;
            FechaFormateada = FechaAsueto.ToString("yyyy-MM-dd 00:00:0");

            //TimeSpan TS_Horas = new TimeSpan(Hora, Minutos, Segundos);//Variable con el horario 19:00 hrs
            NuevaChecada = Convert.ToDateTime(Convert.ToDateTime(FechaFormateada));//Establece fecha de la checada con la hora de ts
            //**************************************************************************************************************************************************/
            if (Turno.Equals("NOCHE"))
            {
                FechaFormateada = FechaAsueto.ToString("yyyy-MM-dd 19:00:0");
                EntradaT = Convert.ToDateTime(FechaFormateada);

                FechaFormateada = FechaAsueto.AddDays(+1).ToString("yyyy-MM-dd 07:00:0");
                SalidaT = Convert.ToDateTime(FechaFormateada);

                FC.InsertarChecada(EntradaT, Empleados, "ET", Turno, "Error", "Vacaciones");
                FC.InsertarChecada(SalidaT, Empleados, "ST", Turno, "Error", "Vacaciones");
            }
            if (Turno.Equals("TNO2"))
            {
                FechaFormateada = FechaAsueto.ToString("yyyy-MM-dd 19:00:0");
                EntradaT = Convert.ToDateTime(FechaFormateada);

                FechaFormateada = FechaAsueto.AddDays(+1).ToString("yyyy-MM-dd 07:00:0");
                SalidaT = Convert.ToDateTime(FechaFormateada);

                FC.InsertarChecada(EntradaT, Empleados, "ET", Turno, "Error", "Vacaciones");
                FC.InsertarChecada(SalidaT, Empleados, "ST", Turno, "Error", "Vacaciones");
            }
            if (Turno.Equals("TARDE"))
            {
                FechaFormateada = FechaAsueto.ToString("yyyy-MM-dd 15:30:0");
                EntradaT = Convert.ToDateTime(FechaFormateada);

                FechaFormateada = FechaAsueto.ToString("yyyy-MM-dd 23:00:0");
                SalidaT = Convert.ToDateTime(FechaFormateada);

                FC.InsertarChecada(EntradaT, Empleados, "ET", Turno, "Error", "Vacaciones");
                FC.InsertarChecada(SalidaT, Empleados, "ST", Turno, "Error", "Vacaciones");
            }
            if (Turno.Equals("TDIA"))
            {
                FechaFormateada = FechaAsueto.ToString("yyyy-MM-dd 07:00:0");
                EntradaT = Convert.ToDateTime(FechaFormateada);

                FechaFormateada = FechaAsueto.ToString("yyyy-MM-dd 12:00:0");
                SalidaC = Convert.ToDateTime(FechaFormateada);
                FechaFormateada = FechaAsueto.ToString("yyyy-MM-dd 12:30:0");
                EntradaC = Convert.ToDateTime(FechaFormateada);

                FechaFormateada = FechaAsueto.ToString("yyyy-MM-dd 15:30:0");
                SalidaT = Convert.ToDateTime(FechaFormateada);

                FC.InsertarChecada(EntradaT, Empleados, "ET", Turno, "Error", "Vacaciones");
                FC.InsertarChecada(SalidaC, Empleados, "SC", Turno, "Error", "Vacaciones");
                FC.InsertarChecada(EntradaC, Empleados, "EC", Turno, "Error", "Vacaciones");
                FC.InsertarChecada(SalidaT, Empleados, "ST", Turno, "Error", "Vacaciones");
            }
            if (Turno.Equals("TFIJ"))
            {
                FechaFormateada = FechaAsueto.ToString("yyyy-MM-dd 08:00:0");
                EntradaT = Convert.ToDateTime(FechaFormateada);

                FechaFormateada = FechaAsueto.ToString("yyyy-MM-dd 12:00:0");
                SalidaC = Convert.ToDateTime(FechaFormateada);
                FechaFormateada = FechaAsueto.ToString("yyyy-MM-dd 12:30:0");
                EntradaC = Convert.ToDateTime(FechaFormateada);

                FechaFormateada = FechaAsueto.ToString("yyyy-MM-dd 17:00:0");
                SalidaT = Convert.ToDateTime(FechaFormateada);

                FC.InsertarChecada(EntradaT, Empleados, "ET", Turno, "Error", "Vacaciones");
                FC.InsertarChecada(SalidaC, Empleados, "SC", Turno, "Error", "Vacaciones");
                FC.InsertarChecada(EntradaC, Empleados, "EC", Turno, "Error", "Vacaciones");
                FC.InsertarChecada(SalidaT, Empleados, "ST", Turno, "Error", "Vacaciones");
            }
        }
        private void InsertarAsueto(String NumEmpleado,String Turno)//Inserta los asuetos
        {
            //MessageBox.Show("Empleado: " + NumEmpleado);
            String TipoChecada = listTipoChecada.Text;
            //******************************************* FORMATO DE CHECADA **********************************************************************************/
            DateTime EntradaT = new DateTime();
            DateTime SalidaT = new DateTime();
            DateTime EntradaC = new DateTime();
            DateTime SalidaC = new DateTime();

            DateTime NuevaFecha;
            DateTime NuevaChecada;
            String FechaFormateada;
            NuevaFecha = Convert.ToDateTime(CalendarioIncidencias.SelectionStart.ToShortDateString().ToString());//Se crea parametro para lanzarlo a la funcion insertarRegistro
            FechaFormateada = NuevaFecha.ToString("yyyy-MM-dd 00:00:0");

            //TimeSpan TS_Horas = new TimeSpan(Hora, Minutos, Segundos);//Variable con el horario 19:00 hrs
            NuevaChecada = Convert.ToDateTime(Convert.ToDateTime(FechaFormateada));//Establece fecha de la checada con la hora de ts
            //**************************************************************************************************************************************************/
            if (Turno.Equals("NOCHE"))
            {
                FechaFormateada = NuevaFecha.ToString("yyyy-MM-dd 19:00:0");
                EntradaT = Convert.ToDateTime(FechaFormateada);

                FechaFormateada = NuevaFecha.AddDays(+1).ToString("yyyy-MM-dd 07:00:0");
                SalidaT = Convert.ToDateTime(FechaFormateada);

                FC.InsertarChecada(EntradaT, NumEmpleado, listTipoChecada.Text.ToString(), Turno, "Error", "Asueto");
                FC.InsertarChecada(SalidaT, NumEmpleado, listTipoChecada.Text.ToString(), Turno, "Error", "Asueto");
            }
            if (Turno.Equals("TNO2"))
            {
                FechaFormateada = NuevaFecha.ToString("yyyy-MM-dd 19:00:0");
                EntradaT = Convert.ToDateTime(FechaFormateada);

                FechaFormateada = NuevaFecha.AddDays(+1).ToString("yyyy-MM-dd 07:00:0");
                SalidaT = Convert.ToDateTime(FechaFormateada);

                FC.InsertarChecada(EntradaT, NumEmpleado, listTipoChecada.Text.ToString(), Turno, "Error", "Asueto");
                FC.InsertarChecada(SalidaT, NumEmpleado, listTipoChecada.Text.ToString(), Turno, "Error", "Asueto");
            }
            if (Turno.Equals("TARDE"))
            {
                FechaFormateada = NuevaFecha.ToString("yyyy-MM-dd 15:30:0");
                EntradaT = Convert.ToDateTime(FechaFormateada);

                FechaFormateada = NuevaFecha.ToString("yyyy-MM-dd 23:00:0");
                SalidaT = Convert.ToDateTime(FechaFormateada);

                FC.InsertarChecada(EntradaT, NumEmpleado, listTipoChecada.Text.ToString(), Turno, "Error", "Asueto");
                FC.InsertarChecada(SalidaT, NumEmpleado, listTipoChecada.Text.ToString(), Turno, "Error", "Asueto");
            }
            if (Turno.Equals("TDIA"))
            {
                FechaFormateada = NuevaFecha.ToString("yyyy-MM-dd 07:00:0");
                EntradaT = Convert.ToDateTime(FechaFormateada);

                FechaFormateada = NuevaFecha.ToString("yyyy-MM-dd 12:00:0");
                SalidaC = Convert.ToDateTime(FechaFormateada);
                FechaFormateada = NuevaFecha.ToString("yyyy-MM-dd 12:30:0");
                EntradaC = Convert.ToDateTime(FechaFormateada);

                FechaFormateada = NuevaFecha.ToString("yyyy-MM-dd 15:30:0");
                SalidaT = Convert.ToDateTime(FechaFormateada);

                FC.InsertarChecada(EntradaT, NumEmpleado, listTipoChecada.Text.ToString(), Turno, "Error", "Asueto");
                FC.InsertarChecada(SalidaC, NumEmpleado, listTipoChecada.Text.ToString(), Turno, "Error", "Asueto");
                FC.InsertarChecada(EntradaC, NumEmpleado, listTipoChecada.Text.ToString(), Turno, "Error", "Asueto");
                FC.InsertarChecada(SalidaT, NumEmpleado, listTipoChecada.Text.ToString(), Turno, "Error", "Asueto");
                MessageBox.Show(SalidaT.ToString());
            }
            if (Turno.Equals("TFIJ"))
            {
                FechaFormateada = NuevaFecha.ToString("yyyy-MM-dd 08:00:0");
                EntradaT = Convert.ToDateTime(FechaFormateada);

                FechaFormateada = NuevaFecha.ToString("yyyy-MM-dd 12:00:0");
                SalidaC = Convert.ToDateTime(FechaFormateada);
                FechaFormateada = NuevaFecha.ToString("yyyy-MM-dd 12:30:0");
                EntradaC = Convert.ToDateTime(FechaFormateada);

                FechaFormateada = NuevaFecha.ToString("yyyy-MM-dd 17:00:0");
                SalidaT = Convert.ToDateTime(FechaFormateada);

                FC.InsertarChecada(EntradaT, NumEmpleado, listTipoChecada.Text.ToString(), Turno, "Error", "Asueto");
                FC.InsertarChecada(SalidaC, NumEmpleado, listTipoChecada.Text.ToString(), Turno, "Error", "Asueto");
                FC.InsertarChecada(EntradaC, NumEmpleado, listTipoChecada.Text.ToString(), Turno, "Error", "Asueto");
                FC.InsertarChecada(SalidaT, NumEmpleado, listTipoChecada.Text.ToString(), Turno, "Error", "Asueto");
            }
        }

        private void ButFestivoIncidencia_Click(object sender, EventArgs e)
        {
            DateTime FechaInicio = Convert.ToDateTime(CalendarioIncidencias.SelectionStart.ToShortDateString().ToString());//Se crea parametro para lanzarlo a la funcion insertarRegistro
            DialogResult result = MessageBox.Show("Asignar asueto: " + FechaInicio.ToShortDateString(), "Salir", MessageBoxButtons.YesNoCancel);
            switch (result)
            {
                case DialogResult.Yes:
                    this.ProcesoAsuetoR();
                    MessageBox.Show("Asueto registrado");
                    break;
                case DialogResult.No:
                    break;
                case DialogResult.Cancel:
                    break;
            }
           
        }

        private void ProcesoAsueto()
        {
            ArrayList Empleado = new ArrayList();
            ArrayList Lunes = new ArrayList();
            ArrayList Martes = new ArrayList();
            ArrayList Miercoles = new ArrayList();
            ArrayList Jueves = new ArrayList();
            ArrayList Viernes = new ArrayList();
            ArrayList Sabado = new ArrayList();
            try//try general
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                query = "SELECT DISTINCT Trabajador FROM Checadas ORDER BY Trabajador";
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

                    FechaInicio = FechaInicioIncidencia.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro
                                                              // FechaInicio=FechaInicio.AddDays(-1);//Se agrega 1 dia más
                    FechaFormateada = FechaInicio.ToString("yyyy-MM-dd 00:00:00:000");

                    FechaFinal = FinFechaIncidencia.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro  
                    FechaFinal = FechaFinal.AddDays(1);
                    FinalFormateada = FechaFinal.ToString("yyyy-MM-dd 00:00:00:000");

                    DataTable dt = c.SemanaTrabajador(Empleado[NumNomina].ToString(), FechaFormateada, FinalFormateada);

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
                    String Turno = dt.Rows[0]["Turno"].ToString();
                    DateTime FechaAsueto = CalendarioIncidencias.SelectionStart;//Selecciona la fecha de asueto                                                               
                    String diaSemana = FechaAsueto.DayOfWeek.ToString();//Convierto la fecha en dia de la semana
                    if ((diaSemana.Equals("Lunes") && Lunes.Count == 0) || (diaSemana.Equals("Monday") && Lunes.Count == 0))
                    {
                        InsertarAsueto(Empleado[NumNomina].ToString(), Turno);
                    }
                    else if ((diaSemana.Equals("Martes") && Martes.Count == 0) || (diaSemana.Equals("Tuesday") && Martes.Count == 0))
                    {
                        InsertarAsueto(Empleado[NumNomina].ToString(), Turno);
                    }
                    else if ((diaSemana.Equals("Miércoles") && Miercoles.Count == 0) || (diaSemana.Equals("Wednesday") && Miercoles.Count == 0))
                    {
                        InsertarAsueto(Empleado[NumNomina].ToString(), Turno);
                    }
                    else if ((diaSemana.Equals("Jueves") && Jueves.Count == 0) || (diaSemana.Equals("Thursday") && Jueves.Count == 0))
                    {
                        InsertarAsueto(Empleado[NumNomina].ToString(), Turno);
                    }
                    else if ((diaSemana.Equals("Viernes") && Viernes.Count == 0) || (diaSemana.Equals("Friday") && Viernes.Count == 0))
                    {
                        InsertarAsueto(Empleado[NumNomina].ToString(), Turno);
                    }
                    else if ((diaSemana.Equals("Sábado") && Sabado.Count == 0) || (diaSemana.Equals("Saturday") && Sabado.Count == 0))
                    {
                        InsertarAsueto(Empleado[NumNomina].ToString(), Turno);
                    }
                }//Fin de for de empleados
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButVerSemanaIncidente_Click(object sender, EventArgs e)
        {
            this.VerSemanaIncidentes();
        }

        private void VerSemanaIncidentes()
        {
            if (txtEmpleadoIncidencia.Text.Equals("#Empleado"))
            {
                MessageBox.Show("No hay empleado a buscar");
            }
            else
            {
                if (gridIncidencias.Columns.Count != 0)//Limpia El Grid
                {
                    gridIncidencias.Rows.Clear();
                    gridIncidencias.Refresh();
                    int i = 0;
                    while (i <= gridIncidencias.Columns.Count)
                    {
                        gridIncidencias.Columns.RemoveAt(0);
                        i = +1;
                    }
                }
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////// CONSULTA SEMANA DE LOS TRABAJADORES /////////////////////////////////////////////////////////
                DateTime FechaInicio, FechaFinal;
                String FechaFormateada, FinalFormateada;

                FechaInicio = FechaInicioIncidencia.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro
                                                          // FechaInicio=FechaInicio.AddDays(-1);//Se agrega 1 dia más
                FechaFormateada = FechaInicio.ToString("yyyy-MM-dd 00:00:00:000");

                FechaFinal = FinFechaIncidencia.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro  
                FechaFinal = FechaFinal.AddDays(1);
                FinalFormateada = FechaFinal.ToString("yyyy-MM-dd 00:00:00:000");

                try
                {
                    cn = new SqlConnection(var.GetConexionDB());
                    cn.Open();
                    //importar la libreria de widws form para abrir una ventana 
                    SqlCommand cmd = new SqlCommand("SemanaEmpleado", cn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@FechaInicio", FechaFormateada);
                    cmd.Parameters.AddWithValue("@FechaFinal", FinalFormateada);
                    cmd.Parameters.AddWithValue("@Empleado", txtEmpleadoIncidencia.Text.ToString());
                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.Load(cmd.ExecuteReader());
                    FC.ObetenrDias(dt);//Obtiene los checadas de cada dia
                    var resultado = FC.ObetenrDias(dt);//Contiene la coleccion de ArrayList de la semana
                    ArrayList Lunes = resultado.Item1;
                    ArrayList Martes = resultado.Item2;
                    ArrayList Miercoles = resultado.Item3;
                    ArrayList Jueves = resultado.Item4;
                    ArrayList Viernes = resultado.Item5;
                    ArrayList Sabado = resultado.Item6;

                    //this.Columnas(gridIncidencias);


                    gridIncidencias.Columns.Add("Trabajador", "Trabajador");
                    gridIncidencias.Columns.Add("Turno", "Turno");
                    gridIncidencias.Columns.Add("Estatus", "Estatus");
                    gridIncidencias.Columns.Add("Lunes", "Lunes " + FechaInicio.ToShortDateString());
                    gridIncidencias.Columns.Add("Martes", "Martes " + FechaInicio.AddDays(1).ToShortDateString());
                    gridIncidencias.Columns.Add("Miércoles", "Miércoles " + FechaInicio.AddDays(2).ToShortDateString());
                    gridIncidencias.Columns.Add("Jueves", "Jueves " + FechaInicio.AddDays(3).ToShortDateString());
                    gridIncidencias.Columns.Add("Viernes", "Viernes " + FechaInicio.AddDays(4).ToShortDateString());
                    gridIncidencias.Columns.Add("Sábado", "Sábado " + FechaInicio.AddDays(5).ToShortDateString());

                    gridIncidencias.Rows.Add();
                    for (int i = 0; i < 6; i++)
                    {
                        if (i >= dt.Rows.Count)
                        {
                            gridIncidencias.Rows[i].Cells["Trabajador"].Value = "-";
                            gridIncidencias.Rows[i].Cells["Turno"].Value = "-";
                            gridIncidencias.Rows[i].Cells["Estatus"].Value = "-";
                        }
                        else
                        {
                            gridIncidencias.Rows[i].Cells["Trabajador"].Value = dt.Rows[i]["Trabajador"].ToString();
                            gridIncidencias.Rows[i].Cells["Turno"].Value = dt.Rows[i]["Turno"].ToString();
                            gridIncidencias.Rows[i].Cells["Estatus"].Value = dt.Rows[i]["Estatus"].ToString();
                        }
                        gridIncidencias.Rows.Add();
                    }
                    for (int elementos = 0; elementos < Lunes.Count; elementos++)
                    {
                        if (Lunes[elementos].ToString() != null)
                        {
                            String Horario = Convert.ToDateTime(Lunes[elementos].ToString()).ToString("HH:mm:ss");
                            gridIncidencias.Rows[elementos].Cells["Lunes"].Value = Horario;//Se crea fila de [Empleado y tipo checada]
                        }
                    }
                    for (int elementos = 0; elementos < Martes.Count; elementos++)
                    {
                        if (Martes[elementos].ToString() != null)
                        {
                            String Horario = Convert.ToDateTime(Martes[elementos].ToString()).ToString("HH:mm:ss");
                            gridIncidencias.Rows[elementos].Cells["Martes"].Value = Horario;//Se crea fila de [Empleado y tipo checada]
                        }
                    }
                    for (int elementos = 0; elementos < Miercoles.Count; elementos++)
                    {
                        if (Miercoles[elementos].ToString() != null)
                        {
                            String Horario = Convert.ToDateTime(Miercoles[elementos].ToString()).ToString("HH:mm:ss");
                            gridIncidencias.Rows[elementos].Cells["Miércoles"].Value = Horario;//Se crea fila de [Empleado y tipo checada]
                        }
                    }
                    for (int elementos = 0; elementos < Jueves.Count; elementos++)
                    {
                        if (Jueves[elementos].ToString() != null)
                        {
                            String Horario = Convert.ToDateTime(Jueves[elementos].ToString()).ToString("HH:mm:ss");
                            gridIncidencias.Rows[elementos].Cells["Jueves"].Value = Horario;//Se crea fila de [Empleado y tipo checada]
                        }
                    }
                    for (int elementos = 0; elementos < Viernes.Count; elementos++)
                    {
                        if (Viernes[elementos].ToString() != null)
                        {
                            String Horario = Convert.ToDateTime(Viernes[elementos].ToString()).ToString("HH:mm:ss");
                            gridIncidencias.Rows[elementos].Cells["Viernes"].Value = Horario;//Se crea fila de [Empleado y tipo checada]
                        }
                    }
                    for (int elementos = 0; elementos < Sabado.Count; elementos++)
                    {
                        if (Sabado[elementos].ToString() != null)
                        {
                            String Horario = Convert.ToDateTime(Sabado[elementos].ToString()).ToString("HH:mm:ss");
                            gridIncidencias.Rows[elementos].Cells["Sábado"].Value = Horario;//Se crea fila de [Empleado y tipo checada]
                        }
                    }
                }
                catch (Exception ex) { MessageBox.Show("Error" + ex); }
                //gridIncidencias.Columns["HorasT"].Visible = false;  //Horas totales
                //gridIncidencias.Columns["Trabajador"].Visible = false;//Numero de nomina
            }
        }

        private void TabControlPrincipal_SelectedIndexChanged(object sender, EventArgs e)
        {
           // MessageBox.Show("Entro el tab");
            if (TabControlPrincipal.SelectedTab.Name.ToString().Equals("Incidencia"))
            {
                if (gridErroresIncidencias.Columns.Count != 0)//Limpia El Grid
                {
                    gridErroresIncidencias.Refresh();
                    int i = 0;
                    while (i <= gridErroresIncidencias.Columns.Count)
                    {
                        gridErroresIncidencias.Columns.RemoveAt(0);
                        i = +1;
                    }
                }
                gridErroresIncidencias.DataSource = RH.MostrarIncidentes();
            }
            else if (TabControlPrincipal.SelectedTab.Name.ToString().Equals("BlackList"))
            {
                this.ConsultarBlackList();
            }
            else if (TabControlPrincipal.SelectedTab.Name.ToString().Equals("Total"))
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                string query = "SELECT * FROM Horas ORDER BY Empleado";

                SqlCommand cmd = new SqlCommand(query, cn);
                DataTable t1 = new DataTable();
                using (SqlDataAdapter a = new SqlDataAdapter(cmd))
                {
                    a.Fill(t1);
                }
                cn.Close();
                gridTotal.DataSource = t1;
                gridTotal.Columns["id"].Visible = false;  //Horas totales
            }
            
        }

        private void GridErroresIncidencias_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (this.gridErroresIncidencias.Columns[e.ColumnIndex].Name == "Empleado")
            {
                int indice = Convert.ToInt32(gridErroresIncidencias.CurrentCell.RowIndex.ToString());//Selecciona la fila a quitar
               // MessageBox.Show(gridErroresIncidencias.Rows[indice].Cells["Empleado"].Value.ToString());
                txtEmpleadoIncidencia.Text = gridErroresIncidencias.Rows[indice].Cells["Empleado"].Value.ToString();
                this.VerSemanaIncidentes();
            }
        }

        private void RdbSemanaCompleta_CheckedChanged(object sender)
        {
            if(rdbSemanaCompleta.Checked == true)
            {
                TLunes.Enabled = false;
                TMartes.Enabled = false;
                Tmiercoles.Enabled = false;
                TJueves.Enabled = false;
                TViernes.Enabled = false;
                TSabado.Enabled = false;
                TSemana.Enabled = true;
            }
        }

        private void RdbSemanaVariada_CheckedChanged(object sender)
        {
            if (rdbSemanaVariada.Checked == true)
            {
                TLunes.Enabled = true;
                TMartes.Enabled = true;
                Tmiercoles.Enabled = true;
                TJueves.Enabled = true;
                TViernes.Enabled = true;
                TSabado.Enabled = true;
                TSemana.Enabled = false;
            }
        }

        private void ButSemanaDeRetardis_Click(object sender, EventArgs e)
        {
            try
            {
                gridSemanaDeRetardos.Rows.Clear();
                DateTime FechaInicio = Convert.ToDateTime(CalendarioDeRetardos.SelectionStart.ToShortDateString().ToString());//Se crea parametro para lanzarlo a la funcion insertarRegistro
                DateTime FechaFinal = Convert.ToDateTime(CalendarioDeRetardos.SelectionEnd.AddDays(1).ToShortDateString().ToString());//Se crea parametro para lanzarlo a la funcion insertarRegistro

                DataTable dt = RH.MostrarSemanasRetrasos(FechaInicio, FechaFinal);
                if (dt.Rows.Count > 0)
                {
                    DataView view = new DataView(dt);
                    DataTable Empleados = view.ToTable(true, "Empleado");
                    String EmpleadoVariable = dt.Rows[0]["Empleado"].ToString();
                    int Contador = 0;
                    for (int x = 0; x < dt.Rows.Count; x++)
                    {
                        String Empleado = dt.Rows[x]["Empleado"].ToString();
                        String Checada = dt.Rows[x]["Checada"].ToString();
                        String Turno = dt.Rows[x]["Turno"].ToString();

                        if (Empleado.Equals(EmpleadoVariable))
                        {
                            Contador++;
                            gridSemanaDeRetardos.Rows.Add(Empleado, Checada, Turno);
                            if (x < dt.Rows.Count - 1)
                            {
                                EmpleadoVariable = dt.Rows[x + 1]["Empleado"].ToString();
                                if (!Empleado.Equals(EmpleadoVariable))
                                {
                                    gridSemanaDeRetardos.Rows[x].Cells["Ret_Total"].Value = Contador;
                                    Contador = 0;
                                }
                            }
                            else if ((x + 1) == dt.Rows.Count)
                            {
                                gridSemanaDeRetardos.Rows[x].Cells["Ret_Total"].Value = Contador;
                                Contador = 0;
                            }
                        }
                        // MessageBox.Show(x.ToString() +  "  "  + dt.Rows.Count.ToString());
                    }//FIN DE FOR
                }
                else
                {
                    MessageBox.Show("Sin Registros");
                }
            }catch(Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            } 
        }

        private void ButExcelRetardos_Click(object sender, EventArgs e)
        {
            RH.ExportarExcelRetornos(gridSemanaDeRetardos);
        }

        private void BurRetardosRH_Click(object sender, EventArgs e)
        {
            RH.ExportToExcel(gridSemanaDeRetardos);
        }

        private void ForPrueba_Click(object sender, EventArgs e)
        {
            for (int x = 0; x < gridSemanaDeRetardos.Rows.Count; x++)
            {
                if(gridSemanaDeRetardos.Rows[x].Cells["Ret_Total"].Value != null)
                {
                    MessageBox.Show("Fila: " + x.ToString() + " Valor: " + gridSemanaDeRetardos.Rows[x].Cells["Ret_Total"].Value.ToString());
                }
            }
        }

        private void GuardarSemanaRetardos_Click(object sender, EventArgs e)
        {
            DateTime FechaInicio = Convert.ToDateTime(CalendarioDeRetardos.SelectionStart.ToShortDateString().ToString());//Se crea parametro para lanzarlo a la funcion insertarRegistro
            DateTime FechaFinal = Convert.ToDateTime(CalendarioDeRetardos.SelectionEnd.AddDays(1).ToShortDateString().ToString());//Se crea parametro para lanzarlo a la funcion insertarRegistro
            try
            {
                DialogResult result = MessageBox.Show("Guardar semana del " + FechaInicio.ToShortDateString() + " al " + FechaFinal.ToShortDateString(), "Salir", MessageBoxButtons.YesNoCancel);
                switch (result)
                {
                    case DialogResult.Yes:
                        if (gridSemanaDeRetardos.Rows.Count > 0)
                        {
                            String Empleado, Turno, Total;
                            DateTime Checada;

                            for (int x = 0; x < gridSemanaDeRetardos.Rows.Count; x++)
                            {
                                if (gridSemanaDeRetardos.Rows[x].Cells["Ret_Total"].Value != null)
                                {
                                    Empleado = gridSemanaDeRetardos.Rows[x].Cells["Ret_Empleado"].Value.ToString();
                                    Turno = gridSemanaDeRetardos.Rows[x].Cells["Ret_Turno"].Value.ToString();
                                    Checada = Convert.ToDateTime(gridSemanaDeRetardos.Rows[x].Cells["Ret_Checada"].Value.ToString());
                                    Total = gridSemanaDeRetardos.Rows[x].Cells["Ret_Total"].Value.ToString();
                                    RH.GuardarHistorial_Retardos(Empleado, Turno, Checada, Total);

                                }
                            }
                            MessageBox.Show("Semana registrada");
                        }
                        else
                        {
                            MessageBox.Show("No se tomaron registros");
                        }

                        break;
                    case DialogResult.No:
                        break;
                    case DialogResult.Cancel:
                        break;
                }
            }catch(Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void ButMesRetardos_Click(object sender, EventArgs e)
        {
            gridSemanaDeRetardos.Rows.Clear();
            DateTime FechaInicio = Convert.ToDateTime(CalendarioDeRetardos.SelectionStart.ToShortDateString().ToString());//Se crea parametro para lanzarlo a la funcion insertarRegistro

            DataTable dt = RH.MostrarMesRetardos(FechaInicio);
            if (dt.Rows.Count > 0)
            {
                DataView view = new DataView(dt);
                DataTable Empleados = view.ToTable(true, "Empleado");
                String EmpleadoVariable = dt.Rows[0]["Empleado"].ToString();
                int Contador = 0;
                for (int x = 0; x < dt.Rows.Count; x++)
                {
                    String Empleado = dt.Rows[x]["Empleado"].ToString();
                    String Checada = dt.Rows[x]["Checada"].ToString();
                    String Turno = dt.Rows[x]["Turno"].ToString();

                    if (Empleado.Equals(EmpleadoVariable))
                    {
                        Contador++;
                        gridSemanaDeRetardos.Rows.Add(Empleado, Checada, Turno);
                        if (x < dt.Rows.Count - 1)
                        {
                            EmpleadoVariable = dt.Rows[x + 1]["Empleado"].ToString();
                            if (!Empleado.Equals(EmpleadoVariable))
                            {
                                gridSemanaDeRetardos.Rows[x].Cells["Ret_Total"].Value = Contador;
                                Contador = 0;
                            }
                        }
                        else if ((x + 1) == dt.Rows.Count)
                        {
                            gridSemanaDeRetardos.Rows[x].Cells["Ret_Total"].Value = Contador;
                            Contador = 0;
                        }
                    }
                }//FIN DE FOR
            }
            else
            {
                MessageBox.Show("Sin Registros");
            }
        }

        private void ButSemanaDeRetardos_Click(object sender, EventArgs e)
        {
            try
            {
                gridSemanaDeRetardos.Rows.Clear();
                DateTime FechaInicio = Convert.ToDateTime(CalendarioDeRetardos.SelectionStart.ToShortDateString().ToString());//Se crea parametro para lanzarlo a la funcion insertarRegistro
                DateTime FechaFinal = Convert.ToDateTime(CalendarioDeRetardos.SelectionEnd.AddDays(1).ToShortDateString().ToString());//Se crea parametro para lanzarlo a la funcion insertarRegistro

                DataTable dt = RH.MostrarSemanasRetrasos(FechaInicio, FechaFinal);
                if (dt.Rows.Count > 0)
                {
                    DataView view = new DataView(dt);
                    DataTable Empleados = view.ToTable(true, "Empleado");
                    String EmpleadoVariable = dt.Rows[0]["Empleado"].ToString();
                    int Contador = 0;
                    for (int x = 0; x < dt.Rows.Count; x++)
                    {
                        String Empleado = dt.Rows[x]["Empleado"].ToString();
                        String Checada = dt.Rows[x]["Checada"].ToString();
                        String Turno = dt.Rows[x]["Turno"].ToString();

                        if (Empleado.Equals(EmpleadoVariable))
                        {
                            Contador++;
                            gridSemanaDeRetardos.Rows.Add(Empleado, Checada, Turno);
                            if (x < dt.Rows.Count - 1)
                            {
                                EmpleadoVariable = dt.Rows[x + 1]["Empleado"].ToString();
                                if (!Empleado.Equals(EmpleadoVariable))
                                {
                                    gridSemanaDeRetardos.Rows[x].Cells["Ret_Total"].Value = Contador;
                                    Contador = 0;
                                }
                            }
                            else if ((x + 1) == dt.Rows.Count)
                            {
                                gridSemanaDeRetardos.Rows[x].Cells["Ret_Total"].Value = Contador;
                                Contador = 0;
                            }
                        }
                    }//FIN DE FOR
                }
                else
                {
                    MessageBox.Show("Sin Registros");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void CargarBase_Click(object sender, EventArgs e)
        {
            chekFiltrar.Checked = false;
            chekDoble.Checked = false;
            chekEntrada.Checked = false;
            chekComida.Checked = false;
            chekSalida.Checked = false;
            chekSI.Checked = false;
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
            this.EliminarBlackList();
            this.MostrarSemanaPrincipal();//Muestra la semana
        }

        private void CargarChecadasTemporal()
        {
            BarraProgreso.Maximum = Convert.ToInt32(CargarChecadas.Rows.Count);
            BarraProgreso.BackColor = Color.White;
            BarraProgreso.ProgressColor1 = Color.CornflowerBlue;
            BarraProgreso.ProgressColor2 = Color.DarkRed;
            BarraProgreso.Visible = true;
            BarraProgreso.ForeColor = Color.Cyan;
            for (int i = 0; i < CargarChecadas.Rows.Count; i++)
            {
                if (i % 100 == 0)
                {
                   // Console.WriteLine(i);
                    BarraProgreso.Value = i;
                    BarraProgreso.Update();//Contador de progreso
                    //BarraProgreso.Visible = true;
                }
                try//Elimino la base de datos actual
                {
                    String Trabajador = CargarChecadas.Rows[i]["Trabajador"].ToString();
                    DateTime Checada = Convert.ToDateTime(CargarChecadas.Rows[i]["Checada"].ToString());
                    String TipoChecada = CargarChecadas.Rows[i]["TipoChecada"].ToString();
                    if (!TipoChecada.Equals("NO"))
                    {
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
                }
                try//Elimino la base de datos actual
                {
                    String Trabajador = CargarTurnos.Rows[i]["Trabajador"].ToString();
                    DateTime Fecha = Convert.ToDateTime(CargarTurnos.Rows[i]["Fecha"].ToString());
                    String Turno = CargarTurnos.Rows[i]["Turno"].ToString();
                    /*if (Turno.Equals("NOCHE") || Turno.Equals("SAB") || Turno.Equals("TARDE") || Turno.Equals("TDIA") || Turno.Equals("TFIJ") || Turno.Equals("TNO2"))
                    {*/
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
                    //}
                }
                catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            }
            BarraProgreso.Visible = false;
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

        private void EliminarBlackList()
        {
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            SqlCommand cmd = new SqlCommand("EliminarBlackList", cn);
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

        private void TblChecadaBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.tblChecadaBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.checadorDataSet);
        }

        private void ButVerSemanaPrincipal_Click(object sender, EventArgs e)
        {
            this.MostrarSemanaPrincipal();//Muestra la semana
        }

        public void ImportarExcel(String archivo, String hoja)
        {
            try
            {
                String Tarjeta = "", Nombre = "", Tipo = "",Checada ="";
                DateTime Checadadt = new DateTime();

                DataTable dt = c.gridExcel(archivo, hoja);

                BarraExcel.Maximum = dt.Rows.Count;
                BarraExcel.BackColor = Color.White;
                BarraExcel.ProgressColor1 = Color.CornflowerBlue;
                BarraExcel.ProgressColor2 = Color.DarkRed;
                BarraExcel.Visible = true;
                BarraExcel.ForeColor = Color.Cyan;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    BarraExcel.Value = i;
                    BarraExcel.Update();//Contador de progreso

                    Tarjeta = dt.Rows[i]["ID de Usuario"].ToString();
                    Nombre = dt.Rows[i]["Nombre"].ToString();
                    Checadadt = Convert.ToDateTime(dt.Rows[i]["Tiempo"].ToString());
                    Tipo = dt.Rows[i]["Estado"].ToString();
                    Checada = Checadadt.ToString("yyyy-MM-dd hh:mm");
                    OPB.ImnportarBaños(Tarjeta, Nombre, Checada, Tipo);
                }
                BarraExcel.Visible = false;
            }
            catch (Exception ex) { MessageBox.Show("Toma de Valores de Excel " + ex.ToString()); }
        }

        private void ButDuplicado_Click(object sender, EventArgs e)
        {
            if (GridSemana.Rows.Count > 1)
            {
                for (int i = GridSemana.Rows.Count - 1; i >= 0; i--)
                {
                    GridSemana.Rows.RemoveAt(i);
                }
            }
            GridSemana.Refresh();
            try
            {
                DateTime FechaInicio, FechaFinal;
                String FechaFormateada, FinalFormateada;

                FechaInicio = dateInicio.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro
                FechaFormateada = FechaInicio.ToString("yyyy-MM-dd 00:00:00:000");
                FechaFinal = dateFinal.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro  
                FechaFinal = FechaFinal.AddDays(1);
                FinalFormateada = FechaFinal.ToString("yyyy-MM-dd 00:00:00:000");

                Op.ChecadasDuplicadas(FechaFormateada, FinalFormateada);
            }
            catch (Exception x)
            {
            }
            MessageBox.Show("Fin de checadas duplicadas");
            chekDoble.Checked = true;
        }

        private void TxtEmpleadoIncidencia_Enter(object sender, EventArgs e)
        {
            txtEmpleadoIncidencia.Text = "";
        }

        private void TxtEmpleadoIncidencia_Leave(object sender, EventArgs e)
        {
            if (txtEmpleadoIncidencia.Text.Equals(""))
            {
                txtEmpleadoIncidencia.Text = "#Empleado";
            }
        }

        private void DateInicio_ValueChanged(object sender, EventArgs e)
        {
            FechaInicio_Date.Value = dateInicio.Value;
            FechaInicioIncidencia.Value = dateInicio.Value;
            TimesHorarioInicio.Value = dateInicio.Value;
            InicioComparador.Value = dateInicio.Value;
        }

        private void DateFinal_ValueChanged(object sender, EventArgs e)
        {
            FechaFinal_Date.Value = dateFinal.Value;
            FinFechaIncidencia.Value = dateFinal.Value;
            TimesHorarioFinal.Value = dateFinal.Value;
            FinComparador.Value = dateFinal.Value;
        }

        private void ButAgregar_Click(object sender, EventArgs e)
        {
            if (txtBlackList.Text.Equals(""))
            {
                MessageBox.Show("Registra número de nómina");
            }
            else
            {
                try
                {
                    cn = new SqlConnection(var.GetConexionDB());
                    cn.Open();
                    cmd = new SqlCommand("INSERT INTO BlackList (Trabajador,Planta) VALUES ('" + txtBlackList.Text.ToString() + "','" + listPlantaBlack.Text.ToString() + "')", cn);
                    cmd.ExecuteNonQuery();//para que relice la sentencia sql
                    MessageBox.Show("Lista negra actualizada");
                    cn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("No se conecto" + ex.ToString());
                }
            }
            this.ConsultarBlackList();
        }

        private void GridBlackList_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (this.gridBlackList.Columns[e.ColumnIndex].Name == "Quítar")
            {
                int indice = Convert.ToInt32(gridBlackList.CurrentCell.RowIndex.ToString());//Selecciona la fila a quitar
                String Trabajador = gridBlackList.Rows[indice].Cells["Trabajador"].Value.ToString();
                String Planta = gridBlackList.Rows[indice].Cells["Planta"].Value.ToString();
                this.gridBlackList.Rows.RemoveAt(indice);//Remueve la fila seleccionada

                MessageBox.Show(Trabajador + Planta);
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                cmd = new SqlCommand("DELETE BlackList WHERE Trabajador = '" + Trabajador + "' AND Planta = '" + Planta +"'", cn);
                cmd.ExecuteNonQuery();//para que relice la sentencia sql
                cn.Close();
            }
            this.ConsultarBlackList();
        }

        private void ConsultarBlackList()
        {
            gridBlackList.Rows.Clear();

            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            query = "SELECT * FROM BlackList";

            SqlCommand cmd2 = new SqlCommand(query, cn);
            DataTable dt = new DataTable();
            using (SqlDataAdapter a = new SqlDataAdapter(cmd2))
            {
                a.Fill(dt);
            }
            cn.Close();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                String Trabajador = dt.Rows[i]["Trabajador"].ToString();
                String Planta = dt.Rows[i]["Planta"].ToString();
                gridBlackList.Rows.Add(Trabajador, Planta);
            }
        }

        private void GridSemanaTrabajador_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DateTime valorCelda = Convert.ToDateTime(gridSemanaTrabajador.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
                txtChecadOriginal.Text = valorCelda.ToString();
                String Fecha = valorCelda.ToLongDateString();
                String Hora = valorCelda.ToString("HH:mm:ss");
                String Turno = gridSemanaTrabajador.Rows[e.RowIndex].Cells[1].Value.ToString();
                listTurno.SelectedItem = Turno;
                dateChecadaNueva.Value = Convert.ToDateTime(Fecha);
                txtHoraNueva.Text = Hora;
            }
            catch (Exception ex)
            {
            }
        }

        private void ActualizarChecada_Click(object sender, EventArgs e)
        {
            String Horas = txtHoraNueva.Text;
            int Hora = Convert.ToInt32(Horas.Substring(0, 2));
            int Minutos = Convert.ToInt32(Horas.Substring(3, 2));
            int Segundos = Convert.ToInt32(0/*Horas.Substring(6,2)*/);
            DateTime NuevaFecha;
            DateTime NuevaChecada;
            String FechaFormateada;
            NuevaFecha = dateChecadaNueva.Value;//Se crea parametro para lanzarlo a la funcion insertarRegistro
            FechaFormateada = NuevaFecha.ToString("yyyy-MM-dd 00:00:0");

            TimeSpan TS_Horas = new TimeSpan(Hora, Minutos, Segundos);//Variable con el horario 19:00 hrs
            NuevaChecada = Convert.ToDateTime(Convert.ToDateTime(FechaFormateada) + TS_Horas);//Establece fecha de la checada con la hora de ts
            MessageBox.Show(NuevaChecada.ToString());
            try
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                cmd = new SqlCommand("ActualizaChecadas", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Checada",Convert.ToDateTime(txtChecadOriginal.Text.ToString()));
                cmd.Parameters.AddWithValue("@Trabajador", txtEmpleado.Text.ToString());
                cmd.Parameters.AddWithValue("@ChecadaActualizada", NuevaChecada);
                cmd.ExecuteReader();
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se conecto" + ex.ToString());
            }
            this.VerActualizarChecadas();
        }

        private void ButCalcularSemana_Click(object sender, EventArgs e)
        {
            DateTime SalidaNoche = Convert.ToDateTime("07:00:00");
            DateTime SalidaSAB = Convert.ToDateTime("13:00:00");
            DateTime SalidaTARDE = Convert.ToDateTime("23:00:00");
            DateTime SalidaTDIA = Convert.ToDateTime("15:30:00");
            DateTime SalidaTFIJ = Convert.ToDateTime("17:00:00");
            DateTime SalidaTNO2 = Convert.ToDateTime("19:00:00");

            ArrayList Lunes = new ArrayList();
            ArrayList Martes = new ArrayList();
            ArrayList Miercoles = new ArrayList();
            ArrayList Jueves = new ArrayList();
            ArrayList Viernes = new ArrayList();
            ArrayList Sabado = new ArrayList();
            ArrayList Semana = new ArrayList { "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado" };
            ArrayList Turno = new ArrayList { "NOCHE", "SAB", "TARDE", "TDIA", "TFIJ", "TNO2" };
            ArrayList Empleado = new ArrayList();
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            query = "TRUNCATE TABLE Horas";
            cmd = new SqlCommand(query, cn);
            cmd.ExecuteNonQuery();//para que relice la sentencia sql
            cn.Close();

            /////////////////////////////////////GET COLECCIÓN && SELECCIÓN DE EMPLEADOS POR TURNOS ///////////////////////////////
            for (int T = 0; T < 5; T++)
            {
                switch (Turno[T].ToString())
                {
                    case "NOCHE":
                        sm.TurnoNOCHE(this.ListasFiltradas("NOCHE"));
                        break;
                    case "SAB":
                        sm.TurnoSAB(this.ListasFiltradas("SAB"));
                        break;
                    case "TARDE":
                        sm.TurnoTARDE(this.ListasFiltradas("TARDE"));
                        break;
                    case "TDIA":
                        sm.TurnoTDIA(this.ListasFiltradas("TDIA"));
                        break;
                    case "TFIJ":
                        sm.TurnoTFIJ(this.ListasFiltradas("TFIJ"));
                        break;
                    default:
                        break;
                }
            }//FOR DEL EMPLEADO

            cn.Open();
            query = "SELECT DISTINCT Trabajador FROM Checadas Order by Trabajador ";
            cmd = new SqlCommand(query, cn);
            DataTable t1 = new DataTable();
            using (SqlDataAdapter a = new SqlDataAdapter(cmd))
            {
                a.Fill(t1);
            }
            cn.Close();
            for (int x = 0; x < t1.Rows.Count; x++)//FOR[1] : Agrega los Trabajadores a una lista
            {
                cn.Open();
                //importar la libreria de widws form para abrir una ventana 
                cmd = new SqlCommand("GenerarTotal", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Empleado", t1.Rows[x]["Trabajador"].ToString());
                cmd.ExecuteReader();
                cn.Close();
            }
            MessageBox.Show("Fin de cálculo");
        }

        private void TxtUnificar_Click(object sender, EventArgs e)
        {
            try
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                query = "TRUNCATE TABLE Prenomina";
                SqlCommand cmd = new SqlCommand(query, cn);
                cmd.ExecuteNonQuery();//para que relice la sentencia sql
                cn.Close();

                ArrayList Lunes = new ArrayList();
                ArrayList Martes = new ArrayList();
                ArrayList Miercoles = new ArrayList();
                ArrayList Jueves = new ArrayList();
                ArrayList Viernes = new ArrayList();
                ArrayList Sabado = new ArrayList();
                ArrayList Semana = new ArrayList { "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado" };

                //cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                query = "SELECT distinct C.Trabajador,C.Checada,C.Turno, C.Estatus,C.Evento,C.Estatus, H.Total,H.Lunes,H.Martes,H.Miercoles,H.Jueves,H.Miercoles,H.Jueves,H.Viernes,H.Sabado " +
                        " FROM Checadas C, Horas H WHERE C.Trabajador = H.Empleado order by Checada";

                cmd = new SqlCommand(query, cn);
                DataTable dt = new DataTable();
                using (SqlDataAdapter a = new SqlDataAdapter(cmd))
                {
                    a.Fill(dt);
                }
                cn.Close();
                Lunes.Clear();
                Martes.Clear();
                Miercoles.Clear();
                Jueves.Clear();
                Viernes.Clear();
                Viernes.Clear();
                Sabado.Clear();
                /////////////////////////////////////////////// OBTENER CHECADAS DEL DIA //////////////////////////////////////////////////////
                for (int reg = 0; reg < dt.Rows.Count; reg++)
                {
                    DateTime dateValue = new DateTime();
                    CultureInfo cul = CultureInfo.CurrentCulture;
                    String query = "";
                    String Trabajador = dt.Rows[reg]["Trabajador"].ToString();
                    String Checada = dt.Rows[reg]["Checada"].ToString();
                    String Turno = dt.Rows[reg]["Turno"].ToString();
                    String Estatus = dt.Rows[reg]["Estatus"].ToString();
                    String Evento = dt.Rows[reg]["Evento"].ToString();
                    Double Total = Convert.ToDouble(dt.Rows[reg]["Total"].ToString());
                    String Lun = dt.Rows[reg]["Lunes"].ToString();
                    String Mar = dt.Rows[reg]["Martes"].ToString();
                    String Mier = dt.Rows[reg]["Miercoles"].ToString();
                    String Jue = dt.Rows[reg]["Jueves"].ToString();
                    String Vie = dt.Rows[reg]["Viernes"].ToString();
                    String Sa = dt.Rows[reg]["Sabado"].ToString();

                    // Usa la fecha formateada y calcula el número de la semana
                    int NumeroSemana = cul.Calendar.GetWeekOfYear(dateValue, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
                    dateValue = Convert.ToDateTime(dt.Rows[reg]["Checada"].ToString());//Consulto la checada del indice
                    String dia = dateValue.DayOfWeek.ToString();//Convierto la fecha en dia de la semana
                    switch (dia)
                    {
                        case "Lunes":
                        case "Monday":
                            cn = new SqlConnection(var.GetConexionDB());
                            cn.Open();
                            query = "INSERT Prenomina (Trabajador,Turno,Estatus,Evento,Total,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,TLunes)VALUES('" + Trabajador + "','" + Turno + "','" + Estatus + "','" + Evento + "'," + Total + ",'" + Lun + "','" + Mar + "','" + Mier + "','" + Jue + "','" + Vie + "','" + Sa + "','" + Turno + "')";
                            cmd = new SqlCommand(query, cn);
                            cmd.ExecuteNonQuery();//para que relice la sentencia sql
                            cn.Close();
                            break;
                        case "Martes":
                        case "Tuesday":
                            cn = new SqlConnection(var.GetConexionDB());
                            cn.Open();
                            query = "INSERT Prenomina (Trabajador,Turno,Estatus,Evento,Total,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,TMartes)VALUES('" + Trabajador + "','" + Turno + "','" + Estatus + "','" + Evento + "'," + Total + ",'" + Lun + "','" + Mar + "','" + Mier + "','" + Jue + "','" + Vie + "','" + Sa + "','" + Turno + "')";
                            cmd = new SqlCommand(query, cn);
                            cmd.ExecuteNonQuery();//para que relice la sentencia sql
                            cn.Close();
                            break;
                        case "Miércoles":
                        case "Wednesday":
                            cn = new SqlConnection(var.GetConexionDB());
                            cn.Open();
                            query = "INSERT Prenomina (Trabajador,Turno,Estatus,Evento,Total,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,TMiercoles)VALUES('" + Trabajador + "','" + Turno + "','" + Estatus + "','" + Evento + "'," + Total + ",'" + Lun + "','" + Mar + "','" + Mier + "','" + Jue + "','" + Vie + "','" + Sa + "','" + Turno + "')";
                            cmd = new SqlCommand(query, cn);
                            cmd.ExecuteNonQuery();//para que relice la sentencia sql
                            cn.Close();
                            break;
                        case "Jueves":
                        case "Thursday":
                            cn = new SqlConnection(var.GetConexionDB());
                            cn.Open();
                            query = "INSERT Prenomina (Trabajador,Turno,Estatus,Evento,Total,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,TJueves)VALUES('" + Trabajador + "','" + Turno + "','" + Estatus + "','" + Evento + "'," + Total + ",'" + Lun + "','" + Mar + "','" + Mier + "','" + Jue + "','" + Vie + "','" + Sa + "','" + Turno + "')";
                            cmd = new SqlCommand(query, cn);
                            cmd.ExecuteNonQuery();//para que relice la sentencia sql
                            cn.Close();
                            break;
                        case "Viernes":
                        case "Friday":
                            cn = new SqlConnection(var.GetConexionDB());
                            cn.Open();
                            query = "INSERT Prenomina (Trabajador,Turno,Estatus,Evento,Total,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,TViernes)VALUES('" + Trabajador + "','" + Turno + "','" + Estatus + "','" + Evento + "'," + Total + ",'" + Lun + "','" + Mar + "','" + Mier + "','" + Jue + "','" + Vie + "','" + Sa + "','" + Turno + "')";
                            cmd = new SqlCommand(query, cn);
                            cmd.ExecuteNonQuery();//para que relice la sentencia sql
                            cn.Close();
                            break;
                        case "Sábado":
                        case "Saturday":
                            cn = new SqlConnection(var.GetConexionDB());
                            cn.Open();
                            query = "INSERT Prenomina (Trabajador,Turno,Estatus,Evento,Total,Lunes,Martes,Miercoles,Jueves,Viernes,Sabado,TSabado)VALUES('" + Trabajador + "','" + Turno + "','" + Estatus + "','" + Evento + "'," + Total + ",'" + Lun + "','" + Mar + "','" + Mier + "','" + Jue + "','" + Vie + "','" + Sa + "','" + Turno + "')";
                            cmd = new SqlCommand(query, cn);
                            cmd.ExecuteNonQuery();//para que relice la sentencia sql
                            cn.Close();
                            break;
                    }//FIN SWITCH dias
                }//Fin FOR Checada por dias
                MessageBox.Show("Tabla Checadas y Tabla Horas Trabajadas Unidas");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ButEliminarChecadaIncidencia_Click(object sender, EventArgs e)
        {
            try
            {
                String FechaFormateada = "";
                int columnIndex = gridIncidencias.CurrentCell.ColumnIndex;
                int rowIndex = gridIncidencias.CurrentCell.RowIndex;

                String Turno = gridIncidencias.Rows[gridIncidencias.CurrentRow.Index].Cells["Turno"].Value.ToString();

                if (columnIndex == 3)
                {
                    DateTime FechaInicio = FechaInicioIncidencia.Value;
                    DateTime Horario = Convert.ToDateTime(gridIncidencias.CurrentCell.Value);
                    FechaFormateada = FechaInicio.ToString("dd-MM-yyyy" + " " + Horario.ToString("HH:mm:ss"));
                }
                else if (columnIndex == 4)
                {
                    DateTime FechaInicio = FechaInicioIncidencia.Value.AddDays(1);
                    DateTime Horario = Convert.ToDateTime(gridIncidencias.CurrentCell.Value);
                    FechaFormateada = FechaInicio.ToString("dd-MM-yyyy" + " " + Horario.ToString("HH:mm:ss"));
                }
                if (columnIndex == 5)
                {
                    DateTime FechaInicio = FechaInicioIncidencia.Value.AddDays(2);
                    DateTime Horario = Convert.ToDateTime(gridIncidencias.CurrentCell.Value);
                    FechaFormateada = FechaInicio.ToString("dd-MM-yyyy" + " " + Horario.ToString("HH:mm:ss"));
                }
                if (columnIndex == 6)
                {
                    DateTime FechaInicio = FechaInicioIncidencia.Value.AddDays(3);
                    DateTime Horario = Convert.ToDateTime(gridIncidencias.CurrentCell.Value);
                    FechaFormateada = FechaInicio.ToString("dd-MM-yyyy" + " " + Horario.ToString("HH:mm:ss"));
                }
                if (columnIndex == 7)
                {
                    DateTime FechaInicio = FechaInicioIncidencia.Value.AddDays(4);
                    DateTime Horario = Convert.ToDateTime(gridIncidencias.CurrentCell.Value);
                    FechaFormateada = FechaInicio.ToString("dd-MM-yyyy" + " " + Horario.ToString("HH:mm:ss"));
                }
                if (columnIndex == 8)
                {
                    DateTime FechaInicio = FechaInicioIncidencia.Value.AddDays(5);
                    DateTime Horario = Convert.ToDateTime(gridIncidencias.CurrentCell.Value);
                    FechaFormateada = FechaInicio.ToString("dd-MM-yyyy" + " " + Horario.ToString("HH:mm:ss"));
                }

               //TimeSpan Hora = new TimeSpan(Convert.ToDateTime(gridIncidencias.CurrentCell.Value).Hour);//Variable con el horario 19:00 hrs
                TimeSpan Minutos = new TimeSpan(Convert.ToDateTime(gridIncidencias.CurrentCell.Value).Minute);//Variable con el horario 19:00 hrs
                TimeSpan Segundos = new TimeSpan(Convert.ToDateTime(gridIncidencias.CurrentCell.Value).Second);//Variable con el horario 19:00 hrs

               //TimeSpan Horario = (Convert.ToInt32(Hora), Convert.ToInt32(Minutos), Convert.ToInt32(Segundos])); ;

                FC.EliminarChecada(Convert.ToDateTime(FechaFormateada), txtEmpleadoIncidencia.Text.ToString(), Turno);
                MessageBox.Show("Checada eliminada");
                this.VerSemanaIncidentes();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.ToString());
            }
            
        }

        private void GridIncidencias_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            foreach (DataGridViewColumn column in gridIncidencias.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        private void GridSemanaTrabajador_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            foreach (DataGridViewColumn column in gridSemanaTrabajador.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        private void GridSemanaHorarios_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            foreach (DataGridViewColumn column in GridSemanaHorarios.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        private void ButExcelTotalHoras_Click(object sender, EventArgs e)
        {
            RH.ExportarExcelRetornos(gridTotal);
        }

        private void FechaActual(){
            dateInicio.Value = DateTime.Now;
            dateFinal.Value = DateTime.Now;
            FechaInicio_Date.Value = DateTime.Now;
            FechaFinal_Date.Value = DateTime.Now;
            dateChecadaNueva.Value = DateTime.Now;
            FechaInicioIncidencia.Value = DateTime.Now;
            FinFechaIncidencia.Value = DateTime.Now;
            TimesHorarioInicio.Value = DateTime.Now;
            TimesHorarioFinal.Value = DateTime.Now;
            FinComparador.Value = DateTime.Now;
        }






    }
}