using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;//libreria sql
using System.Data;//libreria sql 
using System.Data.Sql;//libreria sql
using System.Windows.Forms;//libreria para usar los messagebox
using System.Globalization;

namespace Nominas
{
    class SumarHoras
    {
        CadenaConexion var = new CadenaConexion();
        String sql;
        String query;
        SqlConnection cn; // para la conexion a SQL
        SqlCommand cmd;  // para las instrucciones SQL
        SqlDataReader dr; //
        public SumarHoras()
        {
        }

        public void TurnoNOCHE(ArrayList Empleados)
        {
            double Horas = 0;
            double TotalHoras = 0;
            ArrayList Lunes = new ArrayList();
            ArrayList Martes = new ArrayList();
            ArrayList Miercoles = new ArrayList();
            ArrayList Jueves = new ArrayList();
            ArrayList Viernes = new ArrayList();
            ArrayList Sabado = new ArrayList();
            ArrayList Evento = new ArrayList();
            ArrayList Semana = new ArrayList { "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado" };
            ArrayList Turno = new ArrayList { "NOCHE", "SAB", "TARDE", "TDIA", "TFIJ", "TNO2" };
            try
            {
                for(int Emp = 0; Emp < Empleados.Count; Emp++)
                {
                    cn = new SqlConnection(var.GetConexionDB());
                    cn.Open();
                    query = "SELECT * FROM Checadas WHERE Trabajador = '" + Empleados[Emp].ToString() + "' AND Turno = 'NOCHE' Order by Trabajador,Checada ";

                    SqlCommand cmd2 = new SqlCommand(query, cn);
                    DataTable dt = new DataTable();

                    using (SqlDataAdapter a = new SqlDataAdapter(cmd2))
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
                        // Usa la fecha formateada y calcula el número de la semana
                        int NumeroSemana = cul.Calendar.GetWeekOfYear(dateValue, CalendarWeekRule.FirstDay, DayOfWeek.Monday);

                        dateValue = Convert.ToDateTime(dt.Rows[reg]["Checada"].ToString());//Consulto la checada del indice
                        String dia = dateValue.DayOfWeek.ToString();//Convierto la fecha en dia de la semana
                        switch (dia)
                        {
                            case "Lunes":
                            case "Monday":
                                Evento.Add(dt.Rows[reg]["Evento"].ToString());
                                Lunes.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                break;
                            case "Martes":
                            case "Tuesday":
                                Evento.Add(dt.Rows[reg]["Evento"].ToString());
                                Martes.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                break;
                            case "Miércoles":
                            case "Wednesday":
                                Evento.Add(dt.Rows[reg]["Evento"].ToString());
                                Miercoles.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                break;
                            case "Jueves":
                            case "Thursday":
                                Evento.Add(dt.Rows[reg]["Evento"].ToString());
                                Jueves.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                break;
                            case "Viernes":
                            case "Friday":
                                Evento.Add(dt.Rows[reg]["Evento"].ToString());
                                Viernes.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                break;
                            case "Sábado":
                            case "Saturday":
                                Evento.Add(dt.Rows[reg]["Evento"].ToString());
                                Sabado.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                break;
                        }//FIN SWITCH dias
                    }//Fin FOR Checada por dias
                     ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                     /////////////////////////////////////////////// SUMA DE HORAS //////////////////////////////////////////////////////
                     ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    int indice = 0;
                    while (indice < 5)
                    {
                        TotalHoras = 0;
                        if (indice == 0)
                        {
                            Horas = 0;
                            if (Lunes.Count == 0)
                            {
                                Horas = 0;
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                            }
                            else if (Lunes.Count == 1)
                            {
                                var horas = (Convert.ToDateTime(Martes[0]) - Convert.ToDateTime(Lunes[0].ToString())).TotalHours;
                                //MessageBox.Show("Lunes: " + Empleado + " Horas: " + horas);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                            }
                            else if (Lunes.Count == 2)
                            {
                                var horas = (((Convert.ToDateTime(Lunes[1]) - Convert.ToDateTime(Lunes[0].ToString())).TotalHours) + (Convert.ToDateTime(Martes[0]) - Convert.ToDateTime(Lunes[1].ToString())).TotalHours);
                                //MessageBox.Show("Lunes: " + Empleado + "Horas: " + horas);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                            }
                            else if (Lunes.Count == 3)
                            {
                                var horas = (((Convert.ToDateTime(Lunes[1]) - Convert.ToDateTime(Lunes[0].ToString())).TotalHours) + (Convert.ToDateTime(Martes[0]) - Convert.ToDateTime(Lunes[2].ToString())).TotalHours);
                                //MessageBox.Show("Lunes: " + Empleado + "Horas: " + horas);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                            }
                        }
                        else if (indice == 1)
                        {
                            Horas = 0;
                            if (Martes.Count == 0)
                            {
                                Horas = 0;
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                            }
                            else if (Martes.Count == 2)
                            {
                                var horas = (Convert.ToDateTime(Miercoles[0]) - Convert.ToDateTime(Martes[1].ToString())).TotalHours;
                                // MessageBox.Show("Martes: " + Empleado + "Horas: " + horas);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                                // MessageBox.Show(Convert.ToDateTime(Martes[1].ToString()).ToString("tt", CultureInfo.InvariantCulture)); Consultar si es AM O PM
                            }
                            else if (Martes.Count == 3)
                            {
                                var horas = ((Convert.ToDateTime(Martes[2]) - Convert.ToDateTime(Martes[1].ToString())).TotalHours + ((Convert.ToDateTime(Miercoles[0]) - Convert.ToDateTime(Martes[2].ToString())).TotalHours));
                                //MessageBox.Show("Martes: " + Empleado + "Horas: " + horas);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                            }
                        }
                        else if (indice == 2)
                        {
                            //MessageBox.Show("Entro Miercoles: " + Miercoles.Count);
                            Horas = 0;
                            if (Miercoles.Count == 0)
                            {
                                Horas = 0;
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                            }
                            else if (Miercoles.Count == 2)
                            {
                                var horas = (Convert.ToDateTime(Jueves[0]) - Convert.ToDateTime(Miercoles[1].ToString())).TotalHours;
                                //MessageBox.Show("Miercoles: " + Empleado + "Horas: " + horas);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                                // MessageBox.Show(Convert.ToDateTime(Martes[1].ToString()).ToString("tt", CultureInfo.InvariantCulture)); Consultar si es AM O PM
                            }
                            else if (Miercoles.Count == 3)
                            {
                                var horas = ((Convert.ToDateTime(Miercoles[2]) - Convert.ToDateTime(Miercoles[1].ToString())).TotalHours + ((Convert.ToDateTime(Jueves[0]) - Convert.ToDateTime(Miercoles[2].ToString())).TotalHours));
                                //MessageBox.Show("Miercoles: " + Empleados[Emp].ToString() + "Horas: " + horas);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                            }
                        }
                        else if (indice == 3)
                        {
                            //MessageBox.Show("Entro Jueves: " + Jueves.Count);
                            Horas = 0;
                            if (Jueves.Count == 0)
                            {
                                Horas = 0;
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                            }
                            else if (Jueves.Count == 2)
                            {
                                var horas = (Convert.ToDateTime(Viernes[0]) - Convert.ToDateTime(Jueves[1].ToString())).TotalHours;
                                // MessageBox.Show("Jueves: " + Empleado + "Horas: " + horas);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                                //MessageBox.Show("Jueves Noche con 2 DT: " + Jueves.Count + " Total: " + TotalHoras.ToString());
                                // MessageBox.Show(Convert.ToDateTime(Martes[1].ToString()).ToString("tt", CultureInfo.InvariantCulture)); Consultar si es AM O PM
                            }
                            else if (Jueves.Count == 3)
                            {
                                var horas = ((Convert.ToDateTime(Jueves[2]) - Convert.ToDateTime(Jueves[1].ToString())).TotalHours + ((Convert.ToDateTime(Viernes[0]) - Convert.ToDateTime(Jueves[2].ToString())).TotalHours));
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                            }
                        }
                        else if (indice == 4)
                        {
                            //MessageBox.Show("Entro Viernes: " + Viernes.Count);
                            Horas = 0;
                            if (Viernes.Count == 0)
                            {
                                Horas = 0;
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                            }
                            else if (Viernes.Count == 2)
                            {
                                var horas = (Convert.ToDateTime(Sabado[0]) - Convert.ToDateTime(Viernes[1].ToString())).TotalHours;
                                //MessageBox.Show("Viernes: " + Empleado + "Horas: " + horas);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                                // MessageBox.Show(Convert.ToDateTime(Martes[1].ToString()).ToString("tt", CultureInfo.InvariantCulture)); Consultar si es AM O PM
                            }
                            else if (Viernes.Count == 3)
                            {
                                var horas = ((Convert.ToDateTime(Viernes[2]) - Convert.ToDateTime(Viernes[1].ToString())).TotalHours + ((Convert.ToDateTime(Sabado[0]) - Convert.ToDateTime(Viernes[2].ToString())).TotalHours));
                                // MessageBox.Show("Viernes: " + Empleado + "Horas: " + horas);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                            }
                        }
                        indice++;
                    }//WHILE DE LA SEMANA
                    Horas = 0;
                    TotalHoras = 0;
                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Sabado");
                }//FOR LISTA DE EMPLEADOS
            }
            catch(Exception ex) {
                MessageBox.Show("Error: " + ex);
            }
        }//Fin NOCHE

        public void TurnoSAB(ArrayList Empleados)
        {
            double Horas = 0;
            double TotalHoras = 0;
            ArrayList Sabado = new ArrayList();
            ArrayList Semana = new ArrayList { "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado" };
            try
            {
                for (int Emp = 0; Emp < Empleados.Count; Emp++)
                {
                    Sabado.Clear();

                    cn = new SqlConnection(var.GetConexionDB());
                    cn.Open();
                    query = "SELECT * FROM Checadas WHERE Trabajador = '" + Empleados[Emp].ToString() + "' AND Turno = 'SAB' Order by Trabajador,Checada";

                    SqlCommand cmd2 = new SqlCommand(query, cn);
                    DataTable dt = new DataTable();
                    using (SqlDataAdapter a = new SqlDataAdapter(cmd2))
                    {
                        a.Fill(dt);
                    }
                    cn.Close();

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
                            case "Sabado":
                            case "Sábado":
                            case "Saturday":
                                Sabado.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                                break;
                        }//FIN SWITCH dias
                    }//Fin FOR Checada por dias
                     ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                     /////////////////////////////////////////////// SUMA DE HORAS //////////////////////////////////////////////////////
                     ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    Horas = 0;
                    TotalHoras = 0;
                    if (Sabado.Count == 0)
                    {
                        TotalHoras = TotalHoras + 0;
                    }
                    else if (Sabado.Count == 2)
                    {
                        var horas = (Convert.ToDateTime(Sabado[1]) - Convert.ToDateTime(Sabado[0].ToString())).TotalHours;
                        //MessageBox.Show("Sabado: " + Empleado + "Horas: " + horas);
                        TotalHoras = TotalHoras + horas;
                        this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Sabado");
                        // MessageBox.Show(Convert.ToDateTime(Martes[1].ToString()).ToString("tt", CultureInfo.InvariantCulture)); Consultar si es AM O PM
                    }
                    //MessageBox.Show("Trabajador: " + Empleado + " Horas trabajadas: " + TotalHoras.ToString());
                }
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex); }
        }//Fin SAB

        public void TurnoTARDE(ArrayList Empleados)
        {
            DateTime SalidaTARDE = Convert.ToDateTime("23:00:00");
            DateTime EntradaTARDE = Convert.ToDateTime("15:30:00");
            double Horas = 0;
            double TotalHoras = 0;
            ArrayList Lunes = new ArrayList();
            ArrayList Martes = new ArrayList();
            ArrayList Miercoles = new ArrayList();
            ArrayList Jueves = new ArrayList();
            ArrayList Viernes = new ArrayList();
            ArrayList Sabado = new ArrayList();
            ArrayList Semana = new ArrayList { "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado" };
            try
            {
                for (int Emp = 0; Emp < Empleados.Count; Emp++)
                {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                query = "SELECT * FROM Checadas WHERE Trabajador = '" + Empleados[Emp].ToString() + "' AND Turno = 'TARDE' ORDER BY Checada";

                SqlCommand cmd2 = new SqlCommand(query, cn);
                DataTable dt = new DataTable();
                using (SqlDataAdapter a = new SqlDataAdapter(cmd2))
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

                 ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                 /////////////////////////////////////////////// SUMA DE HORAS //////////////////////////////////////////////////////
                 ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    int indice = 0;
                    while (indice < 6)
                    {
                    TotalHoras = 0;
                        if (indice == 0)//=========================================================================== LUNES =========================================================================================================
                        {
                            Horas = 0;
                            if (Lunes.Count == 0)
                            {
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                            }
                            else if (Lunes.Count == 2)
                            {
                                var horas = (Convert.ToDateTime(Lunes[1]) - Convert.ToDateTime(Lunes[0].ToString())).TotalHours;
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                            }
                            else if (Lunes.Count == 3)
                            {
                                TimeSpan ETARDE = new TimeSpan(15, 30, 00);//Hora de entrada HH:mm:ss
                                EntradaTARDE = Convert.ToDateTime(Lunes[0].ToString()).Date + ETARDE;//Establece fecha de la checada con la hora de ts
                                TimeSpan STARDE = new TimeSpan(23, 00, 00);//Hora de entrada HH:mm:ss
                                SalidaTARDE = Convert.ToDateTime(Lunes[0].ToString()).Date + STARDE;//Establece fecha de la checada con la hora de ts
                                DateTime EntradaMartes = Convert.ToDateTime(Martes[0].ToString()).Date + ETARDE;//Establece fecha de la checada con la hora de t
                                if (Convert.ToDateTime(Martes[0]) == EntradaMartes)//NO ESTA CUMPLIENDO LA CONDICIÓN
                                {
                                    var horas = (((Convert.ToDateTime(Lunes[1].ToString()) - Convert.ToDateTime(Lunes[0].ToString())).TotalHours));
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                                }
                                else if (Convert.ToDateTime(Lunes[0].ToString()) < EntradaTARDE.AddHours(-1))//Primera Checada Es Entrada Extra
                                {
                                    var horas = (((Convert.ToDateTime(Lunes[1].ToString()) - Convert.ToDateTime(Lunes[0].ToString())).TotalHours) + ((Convert.ToDateTime(Lunes[2]) - Convert.ToDateTime(Lunes[1].ToString())).TotalHours));
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                                }
                                else if ((Convert.ToDateTime(Lunes[2].ToString()) >= SalidaTARDE) && Convert.ToDateTime(Martes[0].ToString()) < EntradaTARDE)
                                {
                                    var horas = (((Convert.ToDateTime(Lunes[1]) - Convert.ToDateTime(Lunes[0].ToString())).TotalHours) + (Convert.ToDateTime(Martes[0]) - Convert.ToDateTime(Lunes[2].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                                }
                                else if ((Convert.ToDateTime(Lunes[2].ToString()) >= SalidaTARDE) && this.CargarTipo(Martes, Empleados[Emp].ToString()).Equals("SI"))
                                {
                                    var horas = (((Convert.ToDateTime(Lunes[1]) - Convert.ToDateTime(Lunes[0].ToString())).TotalHours) + (Convert.ToDateTime(Martes[0]) - Convert.ToDateTime(Lunes[2].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                                }
                                else if ((Convert.ToDateTime(Lunes[2].ToString()) >= SalidaTARDE) && !this.CargarTipo(Martes, Empleados[Emp].ToString()).Equals("SI"))
                                {
                                    var horas = (((Convert.ToDateTime(Lunes[1]) - Convert.ToDateTime(Lunes[0].ToString())).TotalHours) + (Convert.ToDateTime(Lunes[2]) - Convert.ToDateTime(Lunes[1].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                                }
                            }
                            else if (Lunes.Count == 6)
                            {
                                var horas = (((Convert.ToDateTime(Lunes[1]) - Convert.ToDateTime(Lunes[0].ToString())).TotalHours) + (Convert.ToDateTime(Lunes[3]) - Convert.ToDateTime(Lunes[2].ToString())).TotalHours + (Convert.ToDateTime(Lunes[5]) - Convert.ToDateTime(Lunes[4].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                            }
                        }
                        else if (indice == 1)//=========================================================================== MARTES =========================================================================================================
                        {
                            Horas = 0;
                            if (Martes.Count == 0)
                            {
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                            }
                            else if (Martes.Count == 2)
                            {
                                var horas = (Convert.ToDateTime(Martes[1]) - Convert.ToDateTime(Martes[0].ToString())).TotalHours;
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                            }
                            else if (Martes.Count == 3)
                            {
                                TimeSpan ETARDE = new TimeSpan(15, 30, 00);//Hora de entrada HH:mm:ss
                                EntradaTARDE = Convert.ToDateTime(Martes[0].ToString()).Date + ETARDE;//Establece fecha de la checada con la hora de ts
                                TimeSpan STARDE = new TimeSpan(23, 00, 00);//Hora de entrada HH:mm:ss
                                SalidaTARDE = Convert.ToDateTime(Martes[0].ToString()).Date + STARDE;//Establece fecha de la checada con la hora de ts
                                DateTime EntradaMiercoles = Convert.ToDateTime(Miercoles[0].ToString()).Date + ETARDE;//Establece fecha de la checada con la hora de ts
                                if (Convert.ToDateTime(Miercoles[0].ToString()) == EntradaMiercoles)
                                {
                                    var horas = (((Convert.ToDateTime(Martes[1].ToString()) - Convert.ToDateTime(Martes[0].ToString())).TotalHours));
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                                }
                                else if (Convert.ToDateTime(Martes[0].ToString()) < EntradaTARDE.AddHours(-1))//Primera Checada Es Entrada Extra
                                {
                                    var horas = (((Convert.ToDateTime(Martes[1].ToString()) - Convert.ToDateTime(Martes[0].ToString())).TotalHours) + ((Convert.ToDateTime(Martes[2]) - Convert.ToDateTime(Martes[1].ToString())).TotalHours));
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                                }
                                else if ((Convert.ToDateTime(Martes[2].ToString()) >= SalidaTARDE) && this.CargarTipo(Miercoles, Empleados[Emp].ToString()).Equals("SI"))
                                {
                                    var horas = (((Convert.ToDateTime(Martes[1]) - Convert.ToDateTime(Martes[0].ToString())).TotalHours) + (Convert.ToDateTime(Miercoles[0]) - Convert.ToDateTime(Martes[2].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                                }
                                else if ((Convert.ToDateTime(Martes[2].ToString()) >= SalidaTARDE) && !this.CargarTipo(Miercoles, Empleados[Emp].ToString()).Equals("SI"))
                                {
                                    var horas = (((Convert.ToDateTime(Martes[1]) - Convert.ToDateTime(Martes[0].ToString())).TotalHours) + (Convert.ToDateTime(Martes[2]) - Convert.ToDateTime(Martes[1].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                                }
                            }
                            else if (Martes.Count == 4)
                            {
                                var horas = (((Convert.ToDateTime(Martes[2]) - Convert.ToDateTime(Martes[1].ToString())).TotalHours) + (Convert.ToDateTime(Miercoles[0]) - Convert.ToDateTime(Martes[2].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                            }
                            else if (Martes.Count == 6)
                            {
                                var horas = (((Convert.ToDateTime(Martes[1]) - Convert.ToDateTime(Martes[0].ToString())).TotalHours) + (Convert.ToDateTime(Martes[3]) - Convert.ToDateTime(Martes[2].ToString())).TotalHours + (Convert.ToDateTime(Martes[5]) - Convert.ToDateTime(Martes[4].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                            }
                        }
                        else if (indice == 2)//=========================================================================== Miercoles =========================================================================================================
                        {
                            Horas = 0;
                            if (Miercoles.Count == 0)
                            {
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                            }
                            else if (Miercoles.Count == 2)
                            {
                                var horas = (Convert.ToDateTime(Miercoles[1]) - Convert.ToDateTime(Miercoles[0].ToString())).TotalHours;
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                            }
                            else if (Miercoles.Count == 3)
                            {
                                TimeSpan ETARDE = new TimeSpan(15, 30, 00);//Hora de entrada HH:mm:ss
                                EntradaTARDE = Convert.ToDateTime(Miercoles[0].ToString()).Date + ETARDE;//Establece fecha de la checada con la hora de ts
                                TimeSpan STARDE = new TimeSpan(23, 00, 00);//Hora de entrada HH:mm:ss
                                SalidaTARDE = Convert.ToDateTime(Miercoles[0].ToString()).Date + STARDE;//Establece fecha de la checada con la hora de ts
                                DateTime EntradaJueves = Convert.ToDateTime(Jueves[0].ToString()).Date + ETARDE;//Establece fecha de la checada con la hora de ts
                                if (Convert.ToDateTime(Jueves[0].ToString()) == EntradaJueves)
                                {
                                    var horas = (((Convert.ToDateTime(Miercoles[1].ToString()) - Convert.ToDateTime(Miercoles[0].ToString())).TotalHours));
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                                }
                                if (this.CargarTipo(Miercoles, Empleados[Emp].ToString()).Equals("SI"))
                                {
                                    var horas = ((Convert.ToDateTime(Miercoles[2].ToString()) - Convert.ToDateTime(Miercoles[1].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                                }
                                if (Convert.ToDateTime(Miercoles[0].ToString()) < EntradaTARDE.AddHours(-1))//Primera Checada Es Entrada Extra
                                {
                                    var horas = (((Convert.ToDateTime(Miercoles[1].ToString()) - Convert.ToDateTime(Miercoles[0].ToString())).TotalHours) + ((Convert.ToDateTime(Miercoles[2]) - Convert.ToDateTime(Miercoles[1].ToString())).TotalHours));
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                                }
                                else if ((Convert.ToDateTime(Miercoles[2].ToString()) >= SalidaTARDE) && this.CargarTipo(Jueves, Empleados[Emp].ToString()).Equals("SI"))
                                {
                                    var horas = (((Convert.ToDateTime(Miercoles[1]) - Convert.ToDateTime(Miercoles[0].ToString())).TotalHours) + (Convert.ToDateTime(Jueves[0]) - Convert.ToDateTime(Miercoles[2].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                                }
                                else if ((Convert.ToDateTime(Miercoles[2].ToString()) >= SalidaTARDE) && !this.CargarTipo(Jueves, Empleados[Emp].ToString()).Equals("SI"))
                                {
                                    var horas = (((Convert.ToDateTime(Miercoles[1]) - Convert.ToDateTime(Miercoles[0].ToString())).TotalHours) + (Convert.ToDateTime(Miercoles[2]) - Convert.ToDateTime(Miercoles[1].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                                }
                            }
                            else if (Miercoles.Count == 4 && Jueves.Count > 0)
                            {
                                var horas = (((Convert.ToDateTime(Miercoles[2]) - Convert.ToDateTime(Miercoles[1].ToString())).TotalHours) + (Convert.ToDateTime(Jueves[0]) - Convert.ToDateTime(Miercoles[2].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                            }
                            else if (Miercoles.Count == 6)
                            {
                                var horas = (((Convert.ToDateTime(Miercoles[1]) - Convert.ToDateTime(Miercoles[0].ToString())).TotalHours) + (Convert.ToDateTime(Miercoles[3]) - Convert.ToDateTime(Miercoles[2].ToString())).TotalHours + (Convert.ToDateTime(Miercoles[5]) - Convert.ToDateTime(Miercoles[4].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                            }
                        }
                        else if (indice == 3)//=========================================================================== Jueves =========================================================================================================
                        {
                            Horas = 0;
                            if (Jueves.Count == 0)
                            {
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                            }
                            else if (Jueves.Count == 2)
                            {
                                var horas = (Convert.ToDateTime(Jueves[1].ToString()) - Convert.ToDateTime(Jueves[0].ToString())).TotalHours;
                                
                                TotalHoras = TotalHoras + horas;
                                if (Empleados[Emp].ToString().Equals("76"))
                                {
                                }
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                            }
                            else if (Jueves.Count == 3)
                            {
                                TimeSpan ETARDE = new TimeSpan(15, 30, 00);//Hora de entrada HH:mm:ss
                                EntradaTARDE = Convert.ToDateTime(Jueves[0].ToString()).Date + ETARDE;//Establece fecha de la checada con la hora de ts
                                TimeSpan STARDE = new TimeSpan(23, 00, 00);//Hora de entrada HH:mm:ss
                                SalidaTARDE = Convert.ToDateTime(Jueves[0].ToString()).Date + STARDE;//Establece fecha de la checada con la hora de ts
                                DateTime EntradaViernes = Convert.ToDateTime(Viernes[0].ToString()).Date + ETARDE;//Establece fecha de la checada con la hora de ts
                                if (Convert.ToDateTime(Viernes[0].ToString()) == EntradaViernes)
                                {
                                    var horas = (((Convert.ToDateTime(Jueves[1].ToString()) - Convert.ToDateTime(Jueves[0].ToString())).TotalHours));
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                                }
                                if (this.CargarTipo(Jueves, Empleados[Emp].ToString()).Equals("SI"))
                                {
                                    var horas = ((Convert.ToDateTime(Miercoles[2].ToString()) - Convert.ToDateTime(Miercoles[1].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                                }
                                if (Convert.ToDateTime(Jueves[0].ToString()) < EntradaTARDE.AddHours(-1))//Primera Checada Es Entrada Extra
                                {
                                    var horas = (((Convert.ToDateTime(Jueves[1].ToString()) - Convert.ToDateTime(Jueves[0].ToString())).TotalHours) + ((Convert.ToDateTime(Jueves[2]) - Convert.ToDateTime(Jueves[1].ToString())).TotalHours));
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                                }
                                else if ((Convert.ToDateTime(Jueves[2].ToString()) >= SalidaTARDE) && this.CargarTipo(Viernes, Empleados[Emp].ToString()).Equals("SI"))
                                {
                                    var horas = (((Convert.ToDateTime(Jueves[1]) - Convert.ToDateTime(Jueves[0].ToString())).TotalHours) + (Convert.ToDateTime(Viernes[0]) - Convert.ToDateTime(Jueves[2].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                                }
                                else if ((Convert.ToDateTime(Jueves[2].ToString()) >= SalidaTARDE) && !this.CargarTipo(Viernes, Empleados[Emp].ToString()).Equals("SI"))
                                {
                                    var horas = (((Convert.ToDateTime(Jueves[1]) - Convert.ToDateTime(Jueves[0].ToString())).TotalHours) + (Convert.ToDateTime(Jueves[2]) - Convert.ToDateTime(Jueves[1].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                                }
                            }
                            else if (Jueves.Count == 4)
                            {
                                var horas=0.0;
                                if (this.CargarTipo(Jueves, Empleados[Emp].ToString()).Equals("SI"))
                                {
                                    horas = (((Convert.ToDateTime(Jueves[2]) - Convert.ToDateTime(Jueves[1].ToString())).TotalHours) + (Convert.ToDateTime(Viernes[0]) - Convert.ToDateTime(Jueves[2].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                                }
                                else
                                {
                                    horas = (((Convert.ToDateTime(Jueves[1]) - Convert.ToDateTime(Jueves[0].ToString())).TotalHours) + (Convert.ToDateTime(Jueves[3]) - Convert.ToDateTime(Jueves[2].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                                }
                            }
                            else if (Jueves.Count == 6)
                            {
                                var horas = (((Convert.ToDateTime(Jueves[1]) - Convert.ToDateTime(Jueves[0].ToString())).TotalHours) + (Convert.ToDateTime(Jueves[3]) - Convert.ToDateTime(Jueves[2].ToString())).TotalHours + (Convert.ToDateTime(Jueves[5]) - Convert.ToDateTime(Jueves[4].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                            }
                        }
                        else if (indice == 4)//=========================================================================== Viernes =========================================================================================================
                        {
                            Horas = 0;
                            if (Viernes.Count == 0)
                            {
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                            }
                            else if (Viernes.Count == 2)
                            {
                                var horas = (Convert.ToDateTime(Viernes[1]) - Convert.ToDateTime(Viernes[0].ToString())).TotalHours;
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                            }
                            else if (Viernes.Count == 3)
                            {
                                TimeSpan ETARDE = new TimeSpan(15, 30, 00);//Hora de entrada HH:mm:ss
                                EntradaTARDE = Convert.ToDateTime(Viernes[0].ToString()).Date + ETARDE;//Establece fecha de la checada con la hora de ts
                                TimeSpan STARDE = new TimeSpan(23, 00, 00);//Hora de entrada HH:mm:ss
                                SalidaTARDE = Convert.ToDateTime(Viernes[0].ToString()).Date + STARDE;//Establece fecha de la checada con la hora de ts
                                DateTime EntradaSabado = Convert.ToDateTime(Sabado[0].ToString()).Date + ETARDE;//Establece fecha de la checada con la hora de ts
                                if (Convert.ToDateTime(Sabado[0].ToString()) == EntradaSabado)
                                {
                                    var horas = (((Convert.ToDateTime(Viernes[1].ToString()) - Convert.ToDateTime(Viernes[0].ToString())).TotalHours));
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                                }
                                if (this.CargarTipo(Viernes, Empleados[Emp].ToString()).Equals("SI"))
                                {
                                    var horas = ((Convert.ToDateTime(Miercoles[2].ToString()) - Convert.ToDateTime(Miercoles[1].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                                }
                                if (Convert.ToDateTime(Viernes[0].ToString()) < EntradaTARDE.AddHours(-1))//Primera Checada Es Entrada Extra
                                {
                                    var horas = (((Convert.ToDateTime(Viernes[1].ToString()) - Convert.ToDateTime(Viernes[0].ToString())).TotalHours) + ((Convert.ToDateTime(Viernes[2]) - Convert.ToDateTime(Viernes[1].ToString())).TotalHours));
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                                }
                                else if ((Convert.ToDateTime(Viernes[2].ToString()) >= SalidaTARDE) && this.CargarTipo(Sabado, Empleados[Emp].ToString()).Equals("SI"))
                                {
                                    var horas = (((Convert.ToDateTime(Viernes[1]) - Convert.ToDateTime(Viernes[0].ToString())).TotalHours) + (Convert.ToDateTime(Sabado[0]) - Convert.ToDateTime(Viernes[2].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                                }
                                else if ((Convert.ToDateTime(Viernes[2].ToString()) >= SalidaTARDE) && !this.CargarTipo(Sabado, Empleados[Emp].ToString()).Equals("SI"))
                                {
                                    var horas = (((Convert.ToDateTime(Viernes[1]) - Convert.ToDateTime(Viernes[0].ToString())).TotalHours) + (Convert.ToDateTime(Viernes[2]) - Convert.ToDateTime(Viernes[1].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                                }
                            }
                            else if (Viernes.Count == 4)
                            {
                                var horas = (((Convert.ToDateTime(Viernes[1]) - Convert.ToDateTime(Viernes[0].ToString())).TotalHours) + (Convert.ToDateTime(Viernes[3]) - Convert.ToDateTime(Viernes[2].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                            }
                            else if (Viernes.Count == 6)
                            {
                                var horas = (((Convert.ToDateTime(Viernes[1]) - Convert.ToDateTime(Viernes[0].ToString())).TotalHours) + (Convert.ToDateTime(Viernes[3]) - Convert.ToDateTime(Viernes[2].ToString())).TotalHours + (Convert.ToDateTime(Viernes[5]) - Convert.ToDateTime(Viernes[4].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                            }
                        }
                        else if (indice == 5)//=========================================================================== Sabado =========================================================================================================
                        {
                            Horas = 0;
                            if (Sabado.Count == 0)
                            {
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Sabado");
                            }
                            else if (Sabado.Count == 2)
                            {
                                var horas = (Convert.ToDateTime(Sabado[1]) - Convert.ToDateTime(Sabado[0].ToString())).TotalHours;
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Sabado");
                            }
                            else if (Sabado.Count == 3)
                            {
                                TimeSpan ETARDE = new TimeSpan(15, 30, 00);//Hora de entrada HH:mm:ss
                                EntradaTARDE = Convert.ToDateTime(Sabado[0].ToString()).Date + ETARDE;//Establece fecha de la checada con la hora de ts
                                TimeSpan STARDE = new TimeSpan(23, 00, 00);//Hora de entrada HH:mm:ss
                                SalidaTARDE = Convert.ToDateTime(Sabado[0].ToString()).Date + STARDE;//Establece fecha de la checada con la hora de ts

                                if (this.CargarTipo(Sabado, Empleados[Emp].ToString()).Equals("ET"))
                                {
                                    var horas = ((Convert.ToDateTime(Sabado[2].ToString()) - Convert.ToDateTime(Sabado[0].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Sabado");
                                }
                                else if (this.CargarTipo(Sabado, Empleados[Emp].ToString()).Equals("SI"))
                                {
                                    var horas = ((Convert.ToDateTime(Sabado[2].ToString()) - Convert.ToDateTime(Sabado[1].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Sabado");
                                }
                                if (Convert.ToDateTime(Sabado[0].ToString()) < EntradaTARDE.AddHours(-1))//Primera Checada Es Entrada Extra
                                {
                                    var horas = (((Convert.ToDateTime(Sabado[1].ToString()) - Convert.ToDateTime(Sabado[0].ToString())).TotalHours) + ((Convert.ToDateTime(Sabado[2]) - Convert.ToDateTime(Sabado[1].ToString())).TotalHours));
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Sabado");
                                }
                            }
                            else if (Sabado.Count == 4)
                            {
                                var horas = (((Convert.ToDateTime(Sabado[1]) - Convert.ToDateTime(Sabado[0].ToString())).TotalHours) + (Convert.ToDateTime(Sabado[3]) - Convert.ToDateTime(Sabado[2].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Sabado");
                            }
                            else if (Sabado.Count == 6)
                            {
                                var horas = (((Convert.ToDateTime(Viernes[1]) - Convert.ToDateTime(Sabado[0].ToString())).TotalHours) + (Convert.ToDateTime(Sabado[3]) - Convert.ToDateTime(Sabado[2].ToString())).TotalHours + (Convert.ToDateTime(Sabado[5]) - Convert.ToDateTime(Sabado[4].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Sabado");
                            }
                        }
                        indice++;
                    }
                }//FOR Lista de empleados
            }
            catch (Exception ex) {
                MessageBox.Show(ex.ToString());
            }
        }


        public void TurnoTDIA(ArrayList Empleados)
        {
            double Horas = 0;
            double TotalHoras = 0;
            ArrayList Lunes = new ArrayList();
            ArrayList Martes = new ArrayList();
            ArrayList Miercoles = new ArrayList();
            ArrayList Jueves = new ArrayList();
            ArrayList Viernes = new ArrayList();
            ArrayList Sabado = new ArrayList();
            ArrayList Semana = new ArrayList { "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado" };
            ArrayList Turno = new ArrayList { "NOCHE", "SAB", "TARDE", "TDIA", "TFIJ", "TNO2" };
            try
            {
                //cn.Close();
                for (int Emp = 0; Emp < Empleados.Count; Emp++)
                {
                    cn = new SqlConnection(var.GetConexionDB());
                    cn.Open();
                    query = "SELECT * FROM Checadas WHERE Trabajador = '" + Empleados[Emp].ToString() + "' AND Turno = 'TDIA' Order by Trabajador,Checada ";

                    SqlCommand cmd2 = new SqlCommand(query, cn);
                    DataTable dt = new DataTable();
                    using (SqlDataAdapter a = new SqlDataAdapter(cmd2))
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
                    }
                        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        /////////////////////////////////////////////// SUMA DE HORAS //////////////////////////////////////////////////////
                        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        int indice = 0;
                        while (indice < 6)
                        {
                        TotalHoras = 0;
                        if (indice == 0)//=========================================================================== LUNES =========================================================================================================
                        {
                            if (Lunes.Count == 0)
                            {
                                TotalHoras = TotalHoras + Horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                            }
                            else if (Lunes.Count == 2)
                            {
                                var horas = ((Convert.ToDateTime(Lunes[1]) - Convert.ToDateTime(Lunes[0].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                            }
                            else if (Lunes.Count == 4)
                            {
                                var horas = ((Convert.ToDateTime(Lunes[1]) - Convert.ToDateTime(Lunes[0].ToString())).TotalHours + (Convert.ToDateTime(Lunes[3]) - Convert.ToDateTime(Lunes[2].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                            }
                            else if (Lunes.Count == 5)
                            {
                                var horas = ((Convert.ToDateTime(Lunes[1]) - Convert.ToDateTime(Lunes[0].ToString())).TotalHours + (Convert.ToDateTime(Lunes[4]) - Convert.ToDateTime(Lunes[2].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                            }
                            else if (Lunes.Count == 6)
                            {
                                var horas = ((Convert.ToDateTime(Lunes[1]) - Convert.ToDateTime(Lunes[0].ToString())).TotalHours + (Convert.ToDateTime(Lunes[3]) - Convert.ToDateTime(Lunes[2].ToString())).TotalHours + (Convert.ToDateTime(Lunes[5]) - Convert.ToDateTime(Lunes[3].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                            }
                        }
                            else if (indice == 1)
                            {
                                if (Martes.Count == 0)
                                {
                                Horas = 0;
                                TotalHoras = TotalHoras + Horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                                }
                            else if (Martes.Count == 2)
                            {
                                var horas = ((Convert.ToDateTime(Martes[1]) - Convert.ToDateTime(Martes[0].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                            }
                            else if (Martes.Count == 4)
                                {
                                    var horas = ((Convert.ToDateTime(Martes[1]) - Convert.ToDateTime(Martes[0].ToString())).TotalHours + (Convert.ToDateTime(Martes[3]) - Convert.ToDateTime(Martes[2].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                                }
                                else if (Martes.Count == 5)
                                {
                                    var horas = ((Convert.ToDateTime(Martes[1]) - Convert.ToDateTime(Martes[0].ToString())).TotalHours + (Convert.ToDateTime(Martes[4]) - Convert.ToDateTime(Martes[2].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                                }
                                else if (Martes.Count == 6)
                                {
                                    var horas = ((Convert.ToDateTime(Martes[1]) - Convert.ToDateTime(Martes[0].ToString())).TotalHours + (Convert.ToDateTime(Martes[3]) - Convert.ToDateTime(Martes[2].ToString())).TotalHours + (Convert.ToDateTime(Martes[5]) - Convert.ToDateTime(Martes[3].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                                }
                            }
                            else if (indice == 2)
                            {
                                if(Miercoles.Count == 0)
                                {
                                Horas = 0;
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                                }
                            else if (Miercoles.Count == 2)
                            {
                                var horas = ((Convert.ToDateTime(Miercoles[1]) - Convert.ToDateTime(Miercoles[0].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                            }
                            else if (Miercoles.Count == 4)
                                {
                                    var horas = ((Convert.ToDateTime(Miercoles[1]) - Convert.ToDateTime(Miercoles[0].ToString())).TotalHours + (Convert.ToDateTime(Miercoles[3]) - Convert.ToDateTime(Miercoles[2].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                                }
                                else if (Miercoles.Count == 5)
                                {
                                    var horas = ((Convert.ToDateTime(Miercoles[1]) - Convert.ToDateTime(Miercoles[0].ToString())).TotalHours + (Convert.ToDateTime(Miercoles[4]) - Convert.ToDateTime(Miercoles[2].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                                }
                                else if (Miercoles.Count == 6)
                                {
                                    var horas = ((Convert.ToDateTime(Miercoles[1]) - Convert.ToDateTime(Miercoles[0].ToString())).TotalHours + (Convert.ToDateTime(Miercoles[3]) - Convert.ToDateTime(Miercoles[2].ToString())).TotalHours + (Convert.ToDateTime(Miercoles[5]) - Convert.ToDateTime(Miercoles[3].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                                }
                            }
                            else if (indice == 3)
                            {  
                                if (Jueves.Count == 0)
                                {
                                Horas = 0;
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                                }
                            if (Jueves.Count == 2)
                            {
                                var horas = ((Convert.ToDateTime(Jueves[1]) - Convert.ToDateTime(Jueves[0].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                            }
                            if (Jueves.Count == 4)
                                {
                                    var horas = ((Convert.ToDateTime(Jueves[1]) - Convert.ToDateTime(Jueves[0].ToString())).TotalHours + (Convert.ToDateTime(Jueves[3]) - Convert.ToDateTime(Jueves[2].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                                }
                                else if (Jueves.Count == 5)
                                {
                                    var horas = ((Convert.ToDateTime(Jueves[1]) - Convert.ToDateTime(Jueves[0].ToString())).TotalHours + (Convert.ToDateTime(Jueves[4]) - Convert.ToDateTime(Jueves[2].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                                }
                                else if (Jueves.Count == 6)
                                {
                                    var horas = ((Convert.ToDateTime(Jueves[1]) - Convert.ToDateTime(Jueves[0].ToString())).TotalHours + (Convert.ToDateTime(Jueves[3]) - Convert.ToDateTime(Jueves[2].ToString())).TotalHours + (Convert.ToDateTime(Jueves[5]) - Convert.ToDateTime(Jueves[3].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                                }
                            }
                            else if (indice == 4)
                            {
                                if (Viernes.Count == 0)
                                {
                                Horas = 0;
                                TotalHoras = TotalHoras + Horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                                }
                            else if (Viernes.Count == 2)
                            {
                                var horas = ((Convert.ToDateTime(Viernes[1]) - Convert.ToDateTime(Viernes[0].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                            }
                            else  if (Viernes.Count == 4)
                                {
                                    var horas = ((Convert.ToDateTime(Viernes[1]) - Convert.ToDateTime(Viernes[0].ToString())).TotalHours + (Convert.ToDateTime(Viernes[3]) - Convert.ToDateTime(Viernes[2].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                                }
                                else if (Viernes.Count == 5)
                                {
                                    var horas = ((Convert.ToDateTime(Viernes[1]) - Convert.ToDateTime(Viernes[0].ToString())).TotalHours + (Convert.ToDateTime(Viernes[4]) - Convert.ToDateTime(Viernes[2].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                                }
                                else if (Viernes.Count == 6)
                                {
                                    var horas = ((Convert.ToDateTime(Viernes[1]) - Convert.ToDateTime(Viernes[0].ToString())).TotalHours + (Convert.ToDateTime(Viernes[3]) - Convert.ToDateTime(Viernes[2].ToString())).TotalHours + (Convert.ToDateTime(Viernes[5]) - Convert.ToDateTime(Viernes[3].ToString())).TotalHours);
                                    TotalHoras = TotalHoras + horas;
                                    this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                                }
                            }
                            else if (indice == 5)
                            {
                            if (Sabado.Count == 0)
                            {
                                Horas = 0;
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Sabado");
                            }
                            else if (Sabado.Count == 2)
                            {
                                var horas = ((Convert.ToDateTime(Sabado[1]) - Convert.ToDateTime(Sabado[0].ToString())).TotalHours );
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Sabado");
                            }
                            else if (Sabado.Count == 4)
                            {
                                var horas = ((Convert.ToDateTime(Sabado[1]) - Convert.ToDateTime(Sabado[0].ToString())).TotalHours + (Convert.ToDateTime(Sabado[3]) - Convert.ToDateTime(Sabado[2].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Sabado");
                            }
                            else if (Sabado.Count == 5)
                            {
                                var horas = ((Convert.ToDateTime(Sabado[1]) - Convert.ToDateTime(Sabado[0].ToString())).TotalHours + (Convert.ToDateTime(Sabado[4]) - Convert.ToDateTime(Sabado[2].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Sabado");
                            }
                            else if (Sabado.Count == 6)
                            {
                                var horas = ((Convert.ToDateTime(Sabado[1]) - Convert.ToDateTime(Sabado[0].ToString())).TotalHours + (Convert.ToDateTime(Sabado[3]) - Convert.ToDateTime(Sabado[2].ToString())).TotalHours + (Convert.ToDateTime(Sabado[5]) - Convert.ToDateTime(Sabado[3].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Sabado");
                            }

                        }
                            indice++;
                        }//find de WHILE
                }//Fin FOR Checada por dias
            }catch(Exception ex) {
                MessageBox.Show("Error Turno TDIA" + ex.Message);
            }
        }

        public void TurnoTFIJ(ArrayList Empleados)
        {
            double Horas = 0;
            double TotalHoras = 0;
            ArrayList Lunes = new ArrayList();
            ArrayList Martes = new ArrayList();
            ArrayList Miercoles = new ArrayList();
            ArrayList Jueves = new ArrayList();
            ArrayList Viernes = new ArrayList();
            ArrayList Sabado = new ArrayList();
            ArrayList Semana = new ArrayList { "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado" };
            ArrayList Turno = new ArrayList { "NOCHE", "SAB", "TARDE", "TDIA", "TFIJ", "TNO2" };
            try
            {
                for (int Emp = 0; Emp < Empleados.Count; Emp++)
                {
                    cn = new SqlConnection(var.GetConexionDB());
                    cn.Open();
                    query = "SELECT * FROM Checadas WHERE Trabajador = '" + Empleados[Emp].ToString() + "' AND Turno = 'TFIJ' Order by Trabajador,Checada ";

                    SqlCommand cmd2 = new SqlCommand(query, cn);
                    DataTable dt = new DataTable();
                    using (SqlDataAdapter a = new SqlDataAdapter(cmd2))
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
                    }

                    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    /////////////////////////////////////////////// SUMA DE HORAS //////////////////////////////////////////////////////
                    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    int indice = 0;
                    while (indice < 6)
                    {
                        TotalHoras = 0;
                        if (indice == 0)//=========================================================================== LUNES =========================================================================================================
                        {
                            Horas = 0;
                            if (Lunes.Count == 0)
                            {
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                            }
                            else if (Lunes.Count == 2)
                            {
                                var horas = ((Convert.ToDateTime(Lunes[1]) - Convert.ToDateTime(Lunes[0].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                            }
                            else if (Lunes.Count == 4)
                            {
                                var horas = ((Convert.ToDateTime(Lunes[1]) - Convert.ToDateTime(Lunes[0].ToString())).TotalHours + (Convert.ToDateTime(Lunes[3]) - Convert.ToDateTime(Lunes[2].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                            }
                            else if (Lunes.Count == 5)
                            {
                                var horas = ((Convert.ToDateTime(Lunes[1]) - Convert.ToDateTime(Lunes[0].ToString())).TotalHours + (Convert.ToDateTime(Lunes[4]) - Convert.ToDateTime(Lunes[2].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                            }
                            else if (Lunes.Count == 6)
                            {
                                var horas = ((Convert.ToDateTime(Lunes[1]) - Convert.ToDateTime(Lunes[0].ToString())).TotalHours + (Convert.ToDateTime(Lunes[3]) - Convert.ToDateTime(Lunes[2].ToString())).TotalHours + (Convert.ToDateTime(Lunes[5]) - Convert.ToDateTime(Lunes[3].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Lunes");
                            }
                        }
                        else if (indice == 1)
                        {
                            Horas = 0;
                            
                            if(Martes.Count == 0)
                            {
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                            }
                            else if (Martes.Count == 2)
                            {
                                var horas = ((Convert.ToDateTime(Martes[1]) - Convert.ToDateTime(Martes[0].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                            }
                            else if (Martes.Count == 4)
                            {
                                var horas = ((Convert.ToDateTime(Martes[1]) - Convert.ToDateTime(Martes[0].ToString())).TotalHours + (Convert.ToDateTime(Martes[3]) - Convert.ToDateTime(Martes[2].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                            }
                            else if (Martes.Count == 5)
                            {
                                var horas = ((Convert.ToDateTime(Martes[1]) - Convert.ToDateTime(Martes[0].ToString())).TotalHours + (Convert.ToDateTime(Martes[4]) - Convert.ToDateTime(Martes[2].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                            }
                            else if (Martes.Count == 6)
                            {
                                var horas = ((Convert.ToDateTime(Martes[1]) - Convert.ToDateTime(Martes[0].ToString())).TotalHours + (Convert.ToDateTime(Martes[3]) - Convert.ToDateTime(Martes[2].ToString())).TotalHours + (Convert.ToDateTime(Martes[5]) - Convert.ToDateTime(Martes[3].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Martes");
                            }
                        }
                        else if (indice == 2)
                        {
                            Horas = 0;
                            if(Miercoles.Count == 0)
                            {
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                            }
                            else if (Miercoles.Count == 2)
                            {
                                var horas = ((Convert.ToDateTime(Miercoles[1]) - Convert.ToDateTime(Miercoles[0].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                            }
                            else if (Miercoles.Count == 4)
                            {
                                var horas = ((Convert.ToDateTime(Miercoles[1]) - Convert.ToDateTime(Miercoles[0].ToString())).TotalHours + (Convert.ToDateTime(Miercoles[3]) - Convert.ToDateTime(Miercoles[2].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                            }
                            else if (Miercoles.Count == 5)
                            {
                                var horas = ((Convert.ToDateTime(Miercoles[1]) - Convert.ToDateTime(Miercoles[0].ToString())).TotalHours + (Convert.ToDateTime(Miercoles[4]) - Convert.ToDateTime(Miercoles[2].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                            }
                            else if (Miercoles.Count == 6)
                            {
                                var horas = ((Convert.ToDateTime(Miercoles[1]) - Convert.ToDateTime(Miercoles[0].ToString())).TotalHours + (Convert.ToDateTime(Miercoles[3]) - Convert.ToDateTime(Miercoles[2].ToString())).TotalHours + (Convert.ToDateTime(Miercoles[5]) - Convert.ToDateTime(Miercoles[3].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Miercoles");
                            }
                        }
                        else if (indice == 3)
                        {
                            Horas = 0;
                           
                            if (Jueves.Count == 0)
                            {
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                            }
                            else if (Jueves.Count == 2)
                            {
                                var horas = ((Convert.ToDateTime(Jueves[1]) - Convert.ToDateTime(Jueves[0].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                            }
                            else if (Jueves.Count == 4)
                            {
                                var horas = ((Convert.ToDateTime(Jueves[1]) - Convert.ToDateTime(Jueves[0].ToString())).TotalHours + (Convert.ToDateTime(Jueves[3]) - Convert.ToDateTime(Jueves[2].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                            }
                            else if (Jueves.Count == 5)
                            {
                                var horas = ((Convert.ToDateTime(Jueves[1]) - Convert.ToDateTime(Jueves[0].ToString())).TotalHours + (Convert.ToDateTime(Jueves[4]) - Convert.ToDateTime(Jueves[2].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                            }
                            else if (Jueves.Count == 6)
                            {
                                var horas = ((Convert.ToDateTime(Jueves[1]) - Convert.ToDateTime(Jueves[0].ToString())).TotalHours + (Convert.ToDateTime(Jueves[3]) - Convert.ToDateTime(Jueves[2].ToString())).TotalHours + (Convert.ToDateTime(Jueves[5]) - Convert.ToDateTime(Jueves[3].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Jueves");
                            }
                        }
                        else if (indice == 4)
                        {
                            Horas = 0;
                            
                            if (Viernes.Count == 0)
                            {
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                            }
                            else if (Viernes.Count == 2)
                            {
                                var horas = ((Convert.ToDateTime(Viernes[1]) - Convert.ToDateTime(Viernes[0].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                            }
                            else if (Viernes.Count == 4)
                            {
                                var horas = ((Convert.ToDateTime(Viernes[1]) - Convert.ToDateTime(Viernes[0].ToString())).TotalHours + (Convert.ToDateTime(Viernes[3]) - Convert.ToDateTime(Viernes[2].ToString())).TotalHours);                                
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                            }
                            else if (Viernes.Count == 5)
                            {
                                var horas = ((Convert.ToDateTime(Viernes[1]) - Convert.ToDateTime(Viernes[0].ToString())).TotalHours + (Convert.ToDateTime(Viernes[4]) - Convert.ToDateTime(Viernes[2].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                            }
                            else if (Viernes.Count == 6)
                            {
                                var horas = ((Convert.ToDateTime(Viernes[1]) - Convert.ToDateTime(Viernes[0].ToString())).TotalHours + (Convert.ToDateTime(Viernes[3]) - Convert.ToDateTime(Viernes[2].ToString())).TotalHours + (Convert.ToDateTime(Viernes[5]) - Convert.ToDateTime(Viernes[3].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Viernes");
                            }
                        }
                        else if (indice == 5)
                        {
                            Horas = 0;
                            if (Sabado.Count == 0)
                            {
                                TotalHoras = 0;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Sabado");
                            }
                            else if (Sabado.Count == 2)
                            {
                                var horas = ((Convert.ToDateTime(Sabado[1]) - Convert.ToDateTime(Sabado[0].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Sabado");
                            }
                            else if (Sabado.Count == 3)
                            {
                                var horas = ((Convert.ToDateTime(Sabado[1]) - Convert.ToDateTime(Sabado[0].ToString())).TotalHours + (Convert.ToDateTime(Sabado[2]) - Convert.ToDateTime(Sabado[1].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Sabado");
                            }
                            else if (Sabado.Count == 4)
                            {
                                var horas = ((Convert.ToDateTime(Sabado[1]) - Convert.ToDateTime(Sabado[0].ToString())).TotalHours + (Convert.ToDateTime(Sabado[3]) - Convert.ToDateTime(Sabado[2].ToString())).TotalHours);
                                TotalHoras = TotalHoras + horas;
                                this.HorasEmpleado(Empleados[Emp].ToString(), TotalHoras, "Sabado");
                            }
                        }
                        indice++;
                    }//find de WHILE
                }//Fin FOR Checada por dias
            }
            catch (Exception ex) {
                MessageBox.Show("Error TFIJ");
            }
        }

        public String CargarTipo(ArrayList ArregloDia, String Empleado)
        {
            String TipoChecada = "*";
            if (ArregloDia.Count > 0)
            {
                DateTime Fecha = Convert.ToDateTime(ArregloDia[0].ToString());
                try {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                String Checada = ArregloDia[0].ToString();

                query = "SELECT * FROM Checadas WHERE Trabajador = '" + Empleado + "' AND Checada = '" + Checada + "'";
                cn.Close();

                cn.Open();
                //importar la libreria de widws form para abrir una ventana 
                SqlCommand cmd = new SqlCommand("TipoChecada", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Fecha", Fecha);
                cmd.Parameters.AddWithValue("@Empleado", Empleado);
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Load(cmd.ExecuteReader());
                String ColItem = dt.Rows[0]["Checada"].ToString();
                TipoChecada = dt.Rows[0]["TipoChecada"].ToString();
                dt.Clear();
                } catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                cn.Close();
            }
            return TipoChecada;
        }

        public void HorasEmpleado(String Empleado, double Horas, String Dia)
        {
            try
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();

                //importar la libreria de widws form para abrir una ventana 
                SqlCommand cmd = new SqlCommand("SumarHoras", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Empleado", Empleado);
                cmd.Parameters.AddWithValue("@Horas", Horas);
                cmd.Parameters.AddWithValue("@Dia", Dia);
                cmd.ExecuteReader();
                cn.Close();
                Dia = "";
                Horas = 0;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }

        private void ProgressStatus(object sender, SqlInfoMessageEventArgs e)
        {
            if (e.Errors.Count > 0)
            {
                string message = e.Errors[0].Message;
                int state = e.Errors[0].State;
                // Set status of the progress bar
            }
        }






    }
}
