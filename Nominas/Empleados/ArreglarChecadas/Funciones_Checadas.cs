using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.Drawing;
using System.Data.SqlClient;//libreria sql
using System.Data;//libreria sql 
using System.Data.Sql;//libreria sql
using System.Windows.Forms;//libreria para usar los messagebox
using System.Data.OleDb;
using System.Globalization;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Http;
using System.Data.SqlClient;//libreria sql
using System.Data;//libreria sql 
using System.Data.Sql;//libreria sql
using System.Windows.Forms;//libreria para usar los messagebox
using System.Data.OleDb;
using System.Collections;

namespace Nominas
{
    class Funciones_Checadas
    {
        SqlConnection cn; // para la conexion a SQL
        SqlCommand cmd;  // para las instrucciones SQL
        SqlDataReader dr; //
        CadenaConexion var = new CadenaConexion();

        public Funciones_Checadas()
        {
        }

        public Tuple<ArrayList, ArrayList, ArrayList, ArrayList, ArrayList, ArrayList> ObetenrDias(DataTable dt)
        {
            ArrayList Lunes = new ArrayList();
            ArrayList Martes = new ArrayList();
            ArrayList Miercoles = new ArrayList();
            ArrayList Jueves = new ArrayList();
            ArrayList Viernes = new ArrayList();
            ArrayList Sabado = new ArrayList();

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

            return Tuple.Create(Lunes, Martes, Miercoles, Jueves, Viernes, Sabado);
        }//FIN OBTENER DIAS

        public void InsertarChecada(DateTime NuevaChecada, String Empleado, String Tipo, String Turno, String Estatus, String Evento)
        {
            try
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                SqlCommand cmd = new SqlCommand("InsertarChecada", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@NuevaChecada", NuevaChecada);
                cmd.Parameters.AddWithValue("@Empleado", Empleado);
                cmd.Parameters.AddWithValue("@Tipo", Tipo);
                cmd.Parameters.AddWithValue("@Turno", Turno);
                cmd.Parameters.AddWithValue("@Estatus", Estatus);
                cmd.Parameters.AddWithValue("@Evento", Evento);
                //cmd.Parameters.AddWithValue("@indice", indice);
                cmd.ExecuteReader();
                cn.Close();
            }
            catch (Exception ex) { MessageBox.Show("No se pudo insertar Checada: " + ex.ToString()); }
        }//Inserta Checadas

        public void InsertarFestivos(DateTime NuevaChecada, String Tipo, String Turno, String Estatus, String Evento)
        {
            try
            {
                //MessageBox.Show("si entra al try: ");
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                // MessageBox.Show("si abre conexion");
                //importar la libreria de widws form para abrir una ventana 
                SqlCommand cmd = new SqlCommand("InsertarChecada", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@NuevaChecada", NuevaChecada);
                cmd.Parameters.AddWithValue("@Tipo", Tipo);
                cmd.Parameters.AddWithValue("@Turno", Turno);
                cmd.Parameters.AddWithValue("@Estatus", Estatus);
                cmd.Parameters.AddWithValue("@Evento", Evento);
                //cmd.Parameters.AddWithValue("@indice", indice);
                cmd.ExecuteReader();
                cn.Close();
            }
            catch (Exception ex) { MessageBox.Show("No se pudo insertar Checada: " + ex.ToString()); }
        }//Inserta Checadas

        public void EliminarChecada(DateTime ChecadaSelect, String Empleado, String Turno)
        {
            try
            {
                //MessageBox.Show("si entra al try: ");
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                // MessageBox.Show("si abre conexion");
                //importar la libreria de widws form para abrir una ventana 
                SqlCommand cmd = new SqlCommand("EliminarChecada", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@ChecadaSelect", ChecadaSelect);
                cmd.Parameters.AddWithValue("@Empleado", Empleado);
                cmd.Parameters.AddWithValue("@Turno", Turno);
                //cmd.Parameters.AddWithValue("@indice", indice);
                cmd.ExecuteReader();
                cn.Close();
            }
            catch (Exception ex) { MessageBox.Show("No se pudo insertar Checada: " + ex.ToString()); }
        }

        public DataTable HorasEmpleado()
        {
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            String query = "SELECT Total FROM Horas";

            SqlCommand cmd = new SqlCommand(query, cn);
            DataTable dt = new DataTable();
            using (SqlDataAdapter a = new SqlDataAdapter(cmd))
            {
                a.Fill(dt);
            }
            cn.Close();
            return dt;
        }

        public void CambiarTurno(String Turno, String Trabajador, String ChecadaInicio, String ChecadaFinal)
        {
            try
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                cmd = new SqlCommand("UPDATE Checadas SET Turno = '" + Turno + "' WHERE Trabajador = '" + Trabajador + "' AND Checada Between '" + ChecadaInicio + "' AND '" + ChecadaFinal + "'", cn);
                cmd.ExecuteNonQuery();//para que relice la sentencia sql
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de importación" + ex.Message.ToString());
            }
        }






































    }
}
