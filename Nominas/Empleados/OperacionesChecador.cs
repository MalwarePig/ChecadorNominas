using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;//libreria sql
using System.Data;//libreria sql 
using System.Data.Sql;//libreria sql
using System.Windows.Forms;//libreria para usar los messagebox
using System.Data.OleDb;
using System.Collections;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Security;

namespace Nominas
{
    class OperacionesChecador
    {
        SqlConnection cn; // para la conexion a SQL
        SqlCommand cmd;  // para las instrucciones SQL
        SqlDataReader dr; //
        CadenaConexion var = new CadenaConexion();
        Funciones_Checadas FC = new Funciones_Checadas();

        public OperacionesChecador()
        {
            //Verificar que se conecte a la base de datos
            try
            {
                cn = new SqlConnection(var.GetConexionDB());
                //cn.Open();
                //importar la libreria de widws form para abrir una ventana 
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se conecto la Base d Datos" + ex.ToString());
            }
        }

        //============================================================= Carga manual de archivo de Excel a un grid ============================================================
        public DataTable gridExcel(String archivo, String hoja)
        {
            DataTable dt = new DataTable();

            //declaramos las variables         
            OleDbConnection conexion = null;
            DataSet dataSet = null;
            OleDbDataAdapter dataAdapter = null;
            String consultaHojaExcel = "Select * from [" + hoja + "$]";
            //esta cadena es para archivos excel 2007 y 2010
            String cadenaConexionArchivoExcel = "provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + archivo + "';Extended Properties=Excel 12.0;";
            if (String.IsNullOrEmpty(hoja))
            {
                MessageBox.Show("No hay una hoja para leer");
            }
            else
            {
                try
                {
                    //Si el usuario escribio el nombre de la hoja se procedera con la busqueda
                    conexion = new OleDbConnection(cadenaConexionArchivoExcel);//creamos la conexion con la hoja de excel
                    conexion.Open(); //abrimos la conexion
                    dataAdapter = new OleDbDataAdapter(consultaHojaExcel, conexion); //traemos los datos de la hoja y las guardamos en un dataAdapter
                    dataSet = new DataSet(); // creamos la instancia del objeto DataSet
                    dataAdapter.Fill(dataSet, hoja);//llenamos el dataset
                    dt = dataSet.Tables[0];
                    conexion.Close();//cerramos la conexion   
                }
                catch (Exception ex)
                {
                    //en caso de haber una excepcion que nos mande un mensaje de error
                    MessageBox.Show("Error, Verificar el archivo o el nombre de la hoja", ex.Message);
                }
            }
            return dt;
        }

        public DataTable tabla(string sql)
        {
            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter();
            DataSet ds = new DataSet();
            da.SelectCommand = new SqlCommand(sql, cn);
            da.Fill(ds);
            dt = ds.Tables[0];
            return dt;
        }

        public void CrearTabla()
        {
            try
            {
                cn.Open();
                String format = "yyyy-MM-dd HH:mm:ss";
           
                cmd = new SqlCommand("TRUNCATE TABLE Checadas", cn);
                cmd.ExecuteNonQuery();//para que relice la sentencia sql
                MessageBox.Show("Termino de crear");
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se conecto" + ex.ToString());
            }
        }

        public void CrearTablaRespaldo()
        {
            try
            {
                cn.Open();
                cmd = new SqlCommand("DROP TABLE TablaOriginal", cn);
                cmd.ExecuteNonQuery();//para que relice la sentencia sql

                cmd = new SqlCommand("SELECT * INTO TablaOriginal FROM Checadas", cn);
                cmd.ExecuteNonQuery();//para que relice la sentencia sql
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se conecto" + ex.ToString());
            }
        }

      
        public void InsertarFalta(String Empleado, String Estatus, DateTime Checada, String Turno)
        {
            try
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                SqlCommand cmd = new SqlCommand("InsertarFalta", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                //cmd.Parameters.AddWithValue("@Checada", Checada);
                cmd.Parameters.AddWithValue("@Trabajador", Empleado);
                cmd.Parameters.AddWithValue("@Estatus", Estatus);
                cmd.Parameters.AddWithValue("@Checada", Checada);
                cmd.Parameters.AddWithValue("@Turno", Turno);
                cmd.ExecuteReader();
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error falta: " + ex.ToString());
            }
        }

        public void ActualizarEstatus(String Empleado, String Estatus)
        {
            try
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                cmd = new SqlCommand("UPDATE Checadas SET Estatus = '" + Estatus + "' WHERE Trabajador = '" + Empleado + "'", cn);
                cmd.ExecuteNonQuery();//para que relice la sentencia sql
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se conecto" + ex.Message.ToString());
            }
        }

        public void ActualizarChecada(String Empleado, DateTime Checada, DateTime EntradaTurno)
        {
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            SqlCommand cmd = new SqlCommand("ActualizaChecadas", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@Checada", Checada);
            cmd.Parameters.AddWithValue("@Trabajador", Empleado);
            cmd.Parameters.AddWithValue("@ChecadaActualizada", EntradaTurno);
            cmd.ExecuteReader();
            cn.Close();
        }

        public void LLegadaTarde(String Empleado, DateTime Checada, DateTime Entrada)
        {
            DateTime Original = Checada;
            int salida = 1;
            if (Empleado.Equals("174"))
            {
                MessageBox.Show("Entro 174");
            }
            while (salida > 0)
            {
                if (salida == 1)
                {
                    if (Checada >= Entrada.AddMinutes(6) && Checada < Entrada.AddMinutes(36))
                    {
                        Checada = Entrada.AddMinutes(30);
                        this.ActualizarChecada(Empleado, Original, Checada);//Actualiza la checada al turno asignado
                        salida = 0;
                    }
                    else
                    {
                        Entrada = Entrada.AddMinutes(30);
                        salida = salida + 1;
                    }
                }
                else
                {
                    if (Checada >= Entrada.AddMinutes(6) && Checada < Entrada.AddMinutes(36))
                    {
                        Checada = Entrada.AddMinutes(30);
                        this.ActualizarChecada(Empleado, Original, Checada);//Actualiza la checada al turno asignado
                        salida = 0;
                    }
                    else
                    {
                        salida = salida + 1;
                        Entrada = Entrada.AddMinutes(30);
                    }
                }
            }//Fin de while
        }

        public void LLegadaTardeAdmin(String Empleado, DateTime Checada, DateTime Entrada)
        {
            MessageBox.Show("Entrada adelantada en: " + Empleado + " Checada: " + Checada.ToString() + " Entrada: " + Entrada.ToString());
            try
            {
                DateTime Original = Checada;
                int salida = 1;
                while (salida > 0)
                {
                    if (salida == 1)
                    {
                        if (Checada >= Entrada.AddMinutes(10) && Checada < Entrada.AddMinutes(15))
                        {
                            Checada = Entrada.AddMinutes(15);
                            this.ActualizarChecada(Empleado, Original, Checada);//Actualiza la checada al turno asignado
                            salida = 0;
                        }
                        else
                        {
                            Entrada = Entrada.AddMinutes(15);
                            salida = salida + 1;
                        }
                    }
                    else
                    {
                        if (Checada >= Entrada.AddMinutes(10) && Checada < Entrada.AddMinutes(15))
                        {
                            Checada = Entrada.AddMinutes(15);
                            this.ActualizarChecada(Empleado, Original, Checada);//Actualiza la checada al turno asignado
                            salida = 0;
                        }
                        else
                        {
                            salida = salida + 1;
                            Entrada = Entrada.AddMinutes(15);
                        }
                    }
                }
            }catch(Exception ex)
            {
                MessageBox.Show("Error en: " + Empleado + " Checada: " + Checada.ToString() + " Entrada: "  + Entrada.ToString()+ " " + ex.Message);
            }
            //Fin de while
        }

        public void ActualizarSalida(String Empleado, DateTime Checada, DateTime SalidaTurno)
        {
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            SqlCommand cmd = new SqlCommand("ActualizaChecadas", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@Checada", Checada);
            cmd.Parameters.AddWithValue("@Trabajador", Empleado);
            cmd.Parameters.AddWithValue("@ChecadaActualizada", SalidaTurno);
            cmd.ExecuteReader();
            cn.Close();
        }

        public void ActualizarChecadaComida(String Empleado,DateTime SalidaComer, DateTime EntradaComida, DateTime EntradaPotencial)
        {
            try
            {
                SqlCommand cmd = new SqlCommand();
                //Quitar segundos a Salida de comida
                DateTime SalidaComidaSinSegundos = SalidaComer; //Variable con la Checada Potencial
                string newSecondsValues = "00";//Valor de segundos
                SalidaComidaSinSegundos = Convert.ToDateTime(SalidaComidaSinSegundos.ToString("dd/MM/yyyy HH:mm:" + newSecondsValues));

               /* cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                cmd = new SqlCommand("UPDATE Checadas SET Checada = '" + SalidaComidaSinSegundos + "' WHERE Trabajador = '" + Empleado + "' AND Checada = '" + SalidaComer + "'", cn);
                cmd.ExecuteNonQuery();//para que relice la sentencia sql
                cn.Close();*/
                //Quitar segundos a regreso de comida
                DateTime EntradaPotencialSinSegundos = EntradaPotencial; //Variable con la Checada Potencial
                string newSecondsValue = "00";//Valor de segundos
                EntradaPotencialSinSegundos = Convert.ToDateTime(EntradaPotencialSinSegundos.ToString("dd/MM/yyyy HH:mm:" + newSecondsValue));
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                //importar la libreria de widws form para abrir una ventana 
                cmd = new SqlCommand("ActualizarChecadaComida", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Checada", EntradaComida);
                cmd.Parameters.AddWithValue("@Trabajador", Empleado);
                cmd.Parameters.AddWithValue("@ChecadaActualizada", EntradaPotencialSinSegundos);
                cmd.Parameters.AddWithValue("@SalidaComidaSinSegundos", SalidaComidaSinSegundos);
                cmd.Parameters.AddWithValue("@SalidaComer", SalidaComer);
                cmd.ExecuteReader();
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void TiempoExtra_Entrada(String Empleado, DateTime Checada, DateTime EntradaTurno)
        {
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            SqlCommand cmd = new SqlCommand("TiempoExtra_Entrada", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@Checada", Checada);
            cmd.Parameters.AddWithValue("@Trabajador", Empleado);
            cmd.Parameters.AddWithValue("@ChecadaActualizada", EntradaTurno);
            cmd.ExecuteReader();
            cn.Close();
        }

        public void CrearTablaExtras()
        {
            try
            {
                cmd = new SqlCommand("DROP TABLE Extras", cn);
                cmd.ExecuteNonQuery();//para que relice la sentencia sql
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se logro eliminar tabla Extras" + ex.Message.ToString());
            }
            try
            {
                cmd = new SqlCommand("CREATE TABLE Extras(ID int IDENTITY(1, 1) PRIMARY KEY,Trabajador varchar(10),Total int default 0,)", cn);
                cmd.ExecuteNonQuery();//para que relice la sentencia sql
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se creo la tabla Extras" + ex.Message.ToString());
            }
        }

        public void ComidaTarde(String Empleado, DateTime SalidaComida, DateTime RegresoComida, DateTime EntradaPotencial, String Turno)
        {
            try
            {
                int salida = 1;
                while (salida > 0)
                { 
                    if (salida == 1)
                    {
                        if ((RegresoComida > SalidaComida && RegresoComida < EntradaPotencial.AddMinutes(6)) && (Turno.Equals("TDIA") || Turno.Equals("TFIJ") || Turno.Equals("TNO2")))
                        {
                            this.ActualizarChecadaComida(Empleado, SalidaComida, RegresoComida, EntradaPotencial);//Actualiza la checada al turno asignado
                            salida = 0;
                        }
                        else
                        {
                            EntradaPotencial = EntradaPotencial.AddMinutes(30);
                            salida = salida + 1;
                        }
                    }
                    else
                    {
                        if ((RegresoComida > SalidaComida && RegresoComida < EntradaPotencial.AddMinutes(6)) && (Turno.Equals("TDIA") || Turno.Equals("TFIJ") || Turno.Equals("TNO2")))
                        {
                            this.ActualizarChecadaComida(Empleado, SalidaComida, RegresoComida, EntradaPotencial);//Actualiza la checada al turno asignado
                            salida = 0;
                        }
                        else
                        {
                            salida = salida + 1;
                            EntradaPotencial = EntradaPotencial.AddMinutes(30);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        public String VerEstatus(String Empleado)
        {
            cn = new SqlConnection(var.GetConexionDB());
            String sql = "SELECT Estatus FROM Checadas WHERE Trabajador='" + Empleado + "'";
            cn.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter();
            DataSet ds = new DataSet();
            da.SelectCommand = new SqlCommand(sql, cn);
            da.Fill(ds);
            dt = ds.Tables[0];
            String Estatus = dt.Rows[0]["Estatus"].ToString();
            cn.Close();
            return Estatus;
        }

        public void AjusteEntrada(String Empleado, DateTime Checada, DateTime SalidaTurno)
        {
            try
            {
                DateTime Original = Checada;
                DateTime SalidaTope = SalidaTurno.AddMinutes(-30);
                int salida = 1;
                while (salida > 0)
                {
                    if (Checada >= SalidaTurno) // CHECADA MAYOR A LA SALIDA
                    {
                        /*if (Empleado.Equals("695"))
                        {
                            MessageBox.Show("1 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString() + "Salida tope: " + SalidaTope.ToString());
                        }*/

                        if ((Checada >= SalidaTurno) && (Checada < SalidaTurno.AddMinutes(6)))
                        {
                            /*if (Empleado.Equals("695"))
                            {
                                MessageBox.Show("2 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                            }*/
                            Checada = SalidaTurno;
                            this.ActualizarChecada(Empleado, Original, Checada);//Actualiza la checada al turno asignado
                            salida = 0;
                        }
                        else if ((Checada >= SalidaTurno.AddMinutes(6)) && (Checada < SalidaTope))
                        {
                           /* if (Empleado.Equals("695"))
                            {
                                MessageBox.Show("3.1.1" + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                            }*/
                            Checada = SalidaTope.AddMinutes(-30);
                            this.ActualizarChecada(Empleado, Original, Checada);//Actualiza la checada al turno asignado
                            salida = 0;
                        }
                        else if ((Checada >= SalidaTurno) && (Checada < SalidaTope.AddMinutes(6)))
                        {
                          /*if (Empleado.Equals("695"))
                            {
                                MessageBox.Show("3.1" + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                            }*/
                            Checada = SalidaTope;
                            this.ActualizarChecada(Empleado, Original, Checada);//Actualiza la checada al turno asignado
                            salida = 0;
                        }
                        else if ((Checada >= SalidaTope) && (Checada < SalidaTope.AddMinutes(6)))
                        {
                            /*if (Empleado.Equals("695"))
                            {
                                MessageBox.Show("3.2 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                            }*/
                            Checada = SalidaTope;
                            this.ActualizarChecada(Empleado, Original, Checada);//Actualiza la checada al turno asignado
                            salida = 0;
                        }
                        else if (Checada > SalidaTope.AddMinutes(6))
                        {
                            /*if (Empleado.Equals("695"))
                            {
                                MessageBox.Show("4 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                            }*/
                            salida = salida + 1;
                            SalidaTope = SalidaTope.AddMinutes(30);
                        }
                        else
                        {
                            /*if (Empleado.Equals("695")) ;
                            {
                                MessageBox.Show("5 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                            }*/
                            salida = salida + 1;
                            SalidaTope = SalidaTope.AddMinutes(30);
                        }
                    }
                    else if (Checada <= SalidaTurno)//Checada es menor o mayo a la entrada
                    {
                        /*if (Empleado.Equals("695"))
                        {
                            MessageBox.Show("6 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                        }*/
                        if (salida == 1)
                        {
                            if ((Checada >= SalidaTurno) && (Checada < SalidaTope))
                            {
                                /*if (Empleado.Equals("695"))
                                {
                                    MessageBox.Show("7 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                                }*/
                                Checada = SalidaTurno;
                                this.ActualizarChecada(Empleado, Original, Checada);//Actualiza la checada al turno asignado
                                salida = 0;
                            }
                            else
                            {
                                /*if (Empleado.Equals("695"))
                                {
                                    MessageBox.Show("8 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                                }*/
                                SalidaTurno = SalidaTurno.AddMinutes(-30);
                                salida = salida + 1;
                            }
                        }
                        else
                        {
                            if ((Checada < SalidaTurno) && (Checada > SalidaTurno.AddMinutes(-30)))
                            {
                                /*if (Empleado.Equals("695"))
                                {
                                    MessageBox.Show("3 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                                }*/
                                Checada = SalidaTurno.AddMinutes(-30);
                                this.ActualizarChecada(Empleado, Original, Checada);//Actualiza la checada al turno asignado
                                salida = 0;
                            }
                            else
                            {
                                salida = salida + 1;
                                SalidaTurno = SalidaTurno.AddMinutes(-30);
                            }
                        }
                    }//else checada < SalidaTurno  */              
                }//while
            }
            catch (Exception ex)
            {
                MessageBox.Show("Empleado: " + Empleado + " Error: " + ex.ToString());
            }
        }

        public void AjusteSalida(String Empleado, DateTime Checada, DateTime SalidaTurno)
        {
            try
            {
                DateTime Original = Checada;
                DateTime SalidaTope = SalidaTurno.AddMinutes(30);
                int salida = 1;
                while (salida > 0)
                {
                    /*if (Empleado.Equals("695"))
                    {
                        MessageBox.Show(Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                    }*/
                    
                    if (Checada >= SalidaTurno) // CHECADA MAYOR A LA SALIDA
                    {
                        /*if (Empleado.Equals("695")) 
                        {
                            MessageBox.Show("1 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString() + "Salida tope: " + SalidaTope.ToString());
                        }*/
                       
                        if ((Checada >= SalidaTurno) && (Checada < SalidaTurno.AddMinutes(6)))
                        {
                            /*if (Empleado.Equals("695")) 
                            {
                                MessageBox.Show("2 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                            }*/
                            Checada = SalidaTurno;
                            this.ActualizarChecada(Empleado, Original, Checada);//Actualiza la checada al turno asignado
                            salida = 0;
                        }
                        else if ((Checada >= SalidaTurno.AddMinutes(6)) && (Checada < SalidaTope))
                        {
                            /*if (Empleado.Equals("695"))
                            {
                                MessageBox.Show("3.1.1" + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                            }*/
                            Checada = SalidaTope.AddMinutes(-30);
                            this.ActualizarChecada(Empleado, Original, Checada);//Actualiza la checada al turno asignado
                            salida = 0;
                        }
                        else if ((Checada >= SalidaTurno) && (Checada < SalidaTope.AddMinutes(6)))
                        {
                            /*if (Empleado.Equals("695"))
                            {
                                MessageBox.Show("3.1" + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                            }*/
                            Checada = SalidaTope;
                            this.ActualizarChecada(Empleado, Original, Checada);//Actualiza la checada al turno asignado
                            salida = 0;
                        }
                        else if ((Checada >= SalidaTope) && (Checada < SalidaTope.AddMinutes(6)))
                        {
                            /*if (Empleado.Equals("695")) 
                            {
                                MessageBox.Show("3.2 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                            }*/
                            Checada = SalidaTope;
                            this.ActualizarChecada(Empleado, Original, Checada);//Actualiza la checada al turno asignado
                            salida = 0;
                        }
                        else if (Checada > SalidaTope.AddMinutes(6))
                        {
                            /*if (Empleado.Equals("695")) 
                            {
                                MessageBox.Show("4 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                            }*/
                            salida = salida + 1;
                            SalidaTope = SalidaTope.AddMinutes(30);
                        }
                        else
                        {
                            /*if (Empleado.Equals("695")) ;
                            {
                                MessageBox.Show("5 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                            }*/
                            salida = salida + 1;
                            SalidaTope = SalidaTope.AddMinutes(30);
                        }
                    }
                    else if (Checada <= SalidaTurno)//Checada es menor o mayo a la entrada
                    {
                        /*if (Empleado.Equals("695")) 
                        {
                            MessageBox.Show("6 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                        }*/
                        if (salida == 1)
                        {
                            if ((Checada >= SalidaTurno) && (Checada < SalidaTope))
                            {
                                /*if (Empleado.Equals("695")) 
                                {
                                   MessageBox.Show("7 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                                }*/
                                Checada = SalidaTurno;
                                this.ActualizarChecada(Empleado, Original, Checada);//Actualiza la checada al turno asignado
                                salida = 0;
                            }
                            else
                            {
                                /*if (Empleado.Equals("695")) 
                                {
                                    MessageBox.Show("8 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                                }*/
                                SalidaTurno = SalidaTurno.AddMinutes(-30);
                                salida = salida + 1;
                            }
                        }
                        else
                        {
                            if ((Checada < SalidaTurno) && (Checada > SalidaTurno.AddMinutes(-30)))
                            {
                                /*if (Empleado.Equals("695")) 
                                {
                                    MessageBox.Show("3 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                                }*/
                                Checada = SalidaTurno.AddMinutes(-30);
                                this.ActualizarChecada(Empleado, Original, Checada);//Actualiza la checada al turno asignado
                                salida = 0;
                            }
                            else
                            {
                                salida = salida + 1;
                                SalidaTurno = SalidaTurno.AddMinutes(-30);
                            }
                        }
                    }//else checada < SalidaTurno  */              
                }//while
            }catch(Exception ex)
            {
                MessageBox.Show("Empleado: " + Empleado + " Error: " + ex.ToString());
            }
        }

        public void AjusteSalidaSI(String Empleado, DateTime Checada, DateTime SalidaTurno)
        {
            try
            {
                DateTime Original = Checada;
                DateTime SalidaTope = SalidaTurno.AddMinutes(30);
                int salida = 1;
                while (salida > 0)
                {
                    if ((Checada >= SalidaTurno) && (Checada < SalidaTope))
                    {
                       /* if (Empleado.Equals("114"))
                        {
                            MessageBox.Show("2 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                        }*/
                        Checada = SalidaTurno;
                        this.ActualizarChecada(Empleado, Original, Checada);//Actualiza la checada al turno asignado
                        salida = 0;
                    }
                    else if ((Checada >= SalidaTurno) && (Checada <= SalidaTope))
                    {
                        /*if (Empleado.Equals("114"))
                        {
                             MessageBox.Show("2 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                        }*/
                        Checada = SalidaTope;
                        this.ActualizarChecada(Empleado, Original, Checada);//Actualiza la checada al turno asignado
                        salida = 0;
                    }
                    else if ((Checada >= SalidaTope) && (Checada < SalidaTope.AddMinutes(30)))
                    {
                        /*if (Empleado.Equals("114"))
                        {
                            MessageBox.Show("3.2 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                        }*/
                        Checada = SalidaTope;
                        this.ActualizarChecada(Empleado, Original, Checada);//Actualiza la checada al turno asignado
                        salida = 0;
                    }
                    else if (Checada > SalidaTope)
                    {
                        /*if (Empleado.Equals("114"))
                        {
                            MessageBox.Show("4 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                        }*/
                        salida = salida + 1;
                        SalidaTope = SalidaTope.AddMinutes(30);
                    }
                    else
                    {
                        /*if (Empleado.Equals("114"))
                        {
                            MessageBox.Show("5 " + Empleado + " Checada: " + Checada.ToString() + " Salida Turno " + SalidaTurno.ToString());
                        }*/
                        salida = salida + 1;
                        SalidaTope = SalidaTope.AddMinutes(30);
                    }
                }//while
            }
            catch (Exception ex)
            {
                // MessageBox.Show("Empleado: " + Empleado + " Error: " + ex.ToString());
            }
        }
        public DataTable Prenomina()
        {
            cn = new SqlConnection(var.GetConexionDB());
            String sql = "SELECT * FROM Checadas";
            cn.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter();
            DataSet ds = new DataSet();
            da.SelectCommand = new SqlCommand(sql, cn);
            da.Fill(ds);
            dt = ds.Tables[0];
            cn.Close();
            return dt;
        }

        public DataTable SemanaTrabajador(String Empleado,String FechaFormateada, String FinalFormateada)
        {
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            //importar la libreria de widws form para abrir una ventana 
            SqlCommand cmd = new SqlCommand("SemanaTrabajador", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@Empleado", Empleado);
            cmd.Parameters.AddWithValue("@FechaInicio", FechaFormateada);
            cmd.Parameters.AddWithValue("@FechaFinal", FinalFormateada);
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Clear();//Limpia antes de llenar
            dt.Load(cmd.ExecuteReader());
            cn.Close();
            return dt;
        }

        public void ChecadasDuplicadas(String Inicio, String Final)
        {
            try
            {
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

            ///////////////////////////////////////////// OBTENER DATOS DE CADA EMPLEADO EN LISTA //////////////////////////////////////////////////////
            ///
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            for (int indiceEmpleados = 0; indiceEmpleados < Empleado.Count; indiceEmpleados++)
            {
                Lunes.Clear();
                Martes.Clear();
                Miercoles.Clear();
                Jueves.Clear();
                Viernes.Clear();
                Viernes.Clear();
                Sabado.Clear();
                String TLunes = "", TMartes = "", TMiercoles = "",TJueves = "", TViernes = "", TSabado = "";
                query = "SELECT * FROM Checadas WHERE Trabajador = '" + Empleado[indiceEmpleados].ToString() + "' Order by Trabajador,Checada ";

                SqlCommand cmd2 = new SqlCommand(query, cn);
                DataTable dt = new DataTable();
                using (SqlDataAdapter a = new SqlDataAdapter(cmd2))
                {
                    a.Fill(dt);
                }

                    //MessageBox.Show(Empleado[indiceEmpleados].ToString());
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
                            TLunes = dt.Rows[reg]["Turno"].ToString();
                            break;
                        case "Martes":
                        case "Tuesday":
                            Martes.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                            TMartes = dt.Rows[reg]["Turno"].ToString();
                            break;
                        case "Miércoles":
                        case "Wednesday":
                            Miercoles.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                            TMiercoles = dt.Rows[reg]["Turno"].ToString();
                            break;
                        case "Jueves":
                        case "Thursday":
                            Jueves.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                            TJueves = dt.Rows[reg]["Turno"].ToString();
                            break;
                        case "Viernes":
                        case "Friday":
                            Viernes.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                            TViernes = dt.Rows[reg]["Turno"].ToString();
                            break;
                        case "Sábado":
                        case "Saturday":
                            Sabado.Add(dt.Rows[reg]["Checada"].ToString());//Agrego registro si es el dia seleccionado
                            TSabado = dt.Rows[reg]["Turno"].ToString();
                            break;
                    }//FIN SWITCH dias
                }//Fin FOR Checada por dias

                    int x = 1;
                    while (x > 0)
                    {
                        switch (x)
                        {
                            case 1:
                                if (Lunes.Count > 1)
                                {
                                    for (int i = 0; i < Lunes.Count; i++)
                                    {
                                        if (i+1 < (Lunes.Count -1 ))
                                        {
                                            if (((Convert.ToDateTime(Lunes[i + 1]) > Convert.ToDateTime(Lunes[i])) && (Convert.ToDateTime(Lunes[i + 1]) < Convert.ToDateTime(Lunes[i]).AddMinutes(5).AddSeconds(59)) || (Convert.ToDateTime(Lunes[i + 1]) == Convert.ToDateTime(Lunes[i]))))
                                            {
                                                Console.WriteLine("Lunes Empleado: " + Empleado[indiceEmpleados].ToString() + " Checada i+1: " + Convert.ToDateTime(Lunes[i + 1].ToString()).ToString() + " Checada i: " + Convert.ToDateTime(Lunes[i].ToString()));
                                                FC.EliminarChecada(Convert.ToDateTime(Lunes[i + 1].ToString()), Empleado[indiceEmpleados].ToString(),TLunes);
                                            }
                                        }
                                    }
                                }
                                break;
                            case 2:
                                if (Martes.Count > 1)
                                {
                                    for (int i = 0; i < Martes.Count; i++)
                                    {
                                        if (i+1 < (Martes.Count - 1))
                                        {
                                            if (((Convert.ToDateTime(Martes[i + 1]) > Convert.ToDateTime(Martes[i])) && (Convert.ToDateTime(Martes[i + 1]) < Convert.ToDateTime(Martes[i]).AddMinutes(5).AddSeconds(59)) || (Convert.ToDateTime(Martes[i + 1]) == Convert.ToDateTime(Martes[i]))))
                                            {
                                                Console.WriteLine("Martes Empleado: " + Empleado[indiceEmpleados].ToString() + " Checada i+1: " + Convert.ToDateTime(Martes[i + 1].ToString()).ToString() + " Checada i: " + Convert.ToDateTime(Martes[i].ToString()));

                                                FC.EliminarChecada(Convert.ToDateTime(Martes[i + 1].ToString()), Empleado[indiceEmpleados].ToString(), TMartes);
                                            }
                                        }
                                    }
                                }
                                break;
                            case 3:
                                if (Miercoles.Count > 1)
                                {
                                    for (int i = 0; i < Miercoles.Count; i++)
                                    {
                                        if (i+1 < (Miercoles.Count - 1))
                                        {
                                            if (((Convert.ToDateTime(Miercoles[i + 1]) > Convert.ToDateTime(Miercoles[i])) && (Convert.ToDateTime(Miercoles[i + 1]) < Convert.ToDateTime(Miercoles[i]).AddMinutes(5).AddSeconds(59)) || (Convert.ToDateTime(Miercoles[i + 1]) == Convert.ToDateTime(Miercoles[i]))))
                                            {
                                                Console.WriteLine(" Miercoles Empleado: " + Empleado[indiceEmpleados].ToString() + " Checada i+1: " + Convert.ToDateTime(Miercoles[i + 1].ToString()).ToString() + " Checada i: " + Convert.ToDateTime(Miercoles[i].ToString()));

                                                FC.EliminarChecada(Convert.ToDateTime(Miercoles[i + 1].ToString()), Empleado[indiceEmpleados].ToString(), TMiercoles);
                                            }
                                        }
                                    }
                                }
                                break;
                            case 4:
                                if (Jueves.Count > 1)
                                {
                                    for (int i = 0; i < Jueves.Count; i++)
                                    {
                                        if (i+1 < (Jueves.Count - 1))
                                        {
                                            if (((Convert.ToDateTime(Jueves[i + 1]) > Convert.ToDateTime(Jueves[i])) && (Convert.ToDateTime(Jueves[i + 1]) < Convert.ToDateTime(Jueves[i]).AddMinutes(5).AddSeconds(59)) || (Convert.ToDateTime(Jueves[i + 1]) == Convert.ToDateTime(Jueves[i]))))
                                            {
                                                Console.WriteLine("Jueves Empleado: " + Empleado[indiceEmpleados].ToString() + " Checada i+1: " + Convert.ToDateTime(Jueves[i + 1].ToString()).ToString() + " Checada i: " + Convert.ToDateTime(Jueves[i].ToString()));

                                                FC.EliminarChecada(Convert.ToDateTime(Jueves[i + 1].ToString()), Empleado[indiceEmpleados].ToString(), TJueves);
                                            }
                                        }
                                    }
                                }
                                break;
                            case 5:
                                if (Viernes.Count > 1)
                                {
                                    for (int i = 0; i < Viernes.Count; i++)
                                    {
                                        if (i < (Viernes.Count - 1))
                                        {
                                            if (((Convert.ToDateTime(Viernes[i + 1]) > Convert.ToDateTime(Viernes[i])) && (Convert.ToDateTime(Viernes[i + 1]) < Convert.ToDateTime(Viernes[i]).AddMinutes(5).AddSeconds(59)) || (Convert.ToDateTime(Viernes[i + 1]) == Convert.ToDateTime(Viernes[i]))))
                                            {
                                                Console.WriteLine("Viernes Empleado: " + Empleado[indiceEmpleados].ToString() + " Checada i+1: " + Convert.ToDateTime(Viernes[i + 1].ToString()).ToString() + " Checada i: " + Convert.ToDateTime(Viernes[i].ToString()));

                                                FC.EliminarChecada(Convert.ToDateTime(Viernes[i + 1].ToString()), Empleado[indiceEmpleados].ToString(), TViernes);
                                            }
                                        }
                                    }
                                }
                                break;
                            case 6:
                                if (Sabado.Count > 1)
                                {
                                    for (int i = 0; i < Sabado.Count; i++)
                                    {
                                        if (i+1 < (Sabado.Count - 1))
                                        {
                                            if (((Convert.ToDateTime(Sabado[i + 1]) > Convert.ToDateTime(Sabado[i])) && (Convert.ToDateTime(Sabado[i + 1]) < Convert.ToDateTime(Sabado[i]).AddMinutes(5).AddSeconds(59)) || (Convert.ToDateTime(Sabado[i + 1]) == Convert.ToDateTime(Sabado[i]))))
                                            {
                                                Console.WriteLine("Sabado Empleado: " + Empleado[indiceEmpleados].ToString() + " Checada i+1: " + Convert.ToDateTime(Sabado[i + 1].ToString()).ToString() + " Checada i: " + Convert.ToDateTime(Sabado[i].ToString()));

                                                FC.EliminarChecada(Convert.ToDateTime(Sabado[i + 1].ToString()), Empleado[indiceEmpleados].ToString(), TSabado);
                                            }
                                        }
                                    }
                                }
                                break;
                            default:
                                //x = 0;
                                break;
                        }//SWITCH
                        if (x==6)
                        {
                            x = 0;
                        }
                        else
                        {
                            x++;
                        }
                    }//WHILE
                }//For de empleados
            }catch (Exception ex)
            {
               
                MessageBox.Show(ex.ToString());
            }
        }//Funcion duplicados




    }
}