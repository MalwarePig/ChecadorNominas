using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;//libreria sql
using System.Data;//libreria sql 
using System.Data.Sql;//libreria sql
using System.Windows.Forms;//libreria para usar los messagebox
using System.Globalization;
using System.Collections;
using System.Data.OleDb;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Http;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Drawing;

namespace Nominas
{
    class OperacionesRH
    {
        CadenaConexion var = new CadenaConexion();
        String sql;
        String query;
        SqlConnection cn; // para la conexion a SQL
        SqlCommand cmd;  // para las instrucciones SQL
        SqlDataReader dr; //
        OperacionesChecador c = new OperacionesChecador();

        public void RegistrarRetardos()
        {
        }

        public void ListaRetardos(String Empleado, String Checadas, String Turno)
        {
            try
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();

                //importar la libreria de widws form para abrir una ventana 
                SqlCommand cmd = new SqlCommand("SP_ListaRetardos", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Empleado", Empleado);
                cmd.Parameters.AddWithValue("@Checadas", Convert.ToDateTime(Checadas));
                cmd.Parameters.AddWithValue("@Turno", Turno);
                cmd.ExecuteReader();
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void LimpiarListaRetardos()
        {
            try
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                cmd = new SqlCommand("TRUNCATE TABLE ListaRetardos", cn);
                cmd.ExecuteNonQuery();//para que relice la sentencia sql
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se limpio tabla ListaRetardos" + ex.Message.ToString());
            }
        }

        public void ListaDeIncidentes(DateTime Inicio,DateTime Fin)
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
                query = "SELECT DISTINCT Trabajador FROM Checadas WHERE Estatus = 'ERROR' ORDER BY Trabajador";
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

                    String FechaFormateada, FinalFormateada;

                    FechaFormateada = Inicio.ToString("yyyy-MM-dd 00:00:00:000");

                    Fin = Fin.AddDays(1);
                    FinalFormateada = Fin.ToString("yyyy-MM-dd 00:00:00:000");

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

                    if (Lunes.Count == 0 || Martes.Count == 0 || Miercoles.Count == 0 || Jueves.Count == 0 || Viernes.Count == 0 || Sabado.Count == 0)
                    {
                        try
                        {
                            cn = new SqlConnection(var.GetConexionDB());
                            cn.Open();
                            cmd = new SqlCommand("INSERT INTO ListaIncidencias(Empleado)VALUES('" + Empleado[NumNomina].ToString() + "')", cn);
                            cmd.ExecuteNonQuery();//para que relice la sentencia sql
                            cn.Close();
                        }
                        catch (Exception ex)
                        {
                            
                        }
                    }
                }//Fin de for de empleados
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public DataTable MostrarIncidentes()
        {
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            //importar la libreria de widws form para abrir una ventana 
            SqlCommand cmd = new SqlCommand("MostrarIncidentes", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Load(cmd.ExecuteReader());
            cn.Close();
            return dt;
        }

        public DataTable MostrarSemanasRetrasos(DateTime Inicio, DateTime fin)
        {
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            //importar la libreria de widws form para abrir una ventana 
            SqlCommand cmd = new SqlCommand("MostrarSemanasRetrasos", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@Inicio", Inicio);
            cmd.Parameters.AddWithValue("@fin", fin);
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Load(cmd.ExecuteReader());
            cn.Close();
            return dt;
        }


        public void ExportarExcelRetornos(DataGridView grd)
        {
            try
            {
                grd.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
                grd.MultiSelect = true;
                grd.SelectAll();
                DataObject dataObj = grd.GetClipboardContent();
                if (dataObj != null)
                    Clipboard.SetDataObject(dataObj);

                Microsoft.Office.Interop.Excel.Application xlexcel;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlexcel = new Excel.Application();
                xlexcel.Visible = true;
                xlWorkBook = xlexcel.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                Excel.Range rango = (Excel.Range)xlWorkSheet.Cells[1, 1];
                rango.Select();
                xlWorkSheet.PasteSpecial(rango, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                // si la primera celda de la primera columna está vacía, elimino la primera columna
                // esto se puede omitir, pero lo dejo para ver cómo se podrían añadir/eliminar datos a posteriori

                // la primera fila en negrita, centrada y con fondo gris
                Excel.Range fila1 = (Excel.Range)xlWorkSheet.Rows[1];
                fila1.Select();
                fila1.EntireRow.Font.Bold = true;
                fila1.EntireRow.HorizontalAlignment = HorizontalAlignment.Center;
                fila1.EntireRow.Interior.Color = Color.Gray;
                fila1.EntireColumn.AutoFit();
                Excel.Range c1f1 = (Excel.Range)xlWorkSheet.Cells[1, 1];
                if (c1f1.Text == "")
                {
                    Excel.Range columna1 = (Excel.Range)xlWorkSheet.Columns[1];
                    columna1.Select();
                    columna1.Delete();
                }

                // selecciono la primera celda de la primera columna
                Excel.Range c1 = (Excel.Range)xlWorkSheet.Cells[1, 1];
                c1.Select();

                grd.ClearSelection();
                grd.MultiSelect = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }//Fin de ExportarExcel



        public void ExportToExcel(DataGridView grd, string excelFilePath = null)
        {
            String Empleado, Total;
            int Fila = 2;
            grd.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            grd.MultiSelect = true;
            grd.SelectAll();
            DataObject dataObj = grd.GetClipboardContent();
            if (dataObj != null) Clipboard.SetDataObject(dataObj);

            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Excel.Application();
            xlexcel.Visible = true;
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Empleado";//INSERTAR INFORMACION EN CELDAS
            xlWorkSheet.Cells[1, 2] = "Total retardos";//INSERTAR INFORMACION EN CELDAS

            for (int x = 0; x < grd.Rows.Count; x++)
            {
                if (grd.Rows[x].Cells["Ret_Total"].Value != null)
                {
                   Empleado = grd.Rows[x].Cells["Ret_Empleado"].Value.ToString();
                   Total = grd.Rows[x].Cells["Ret_Total"].Value.ToString();
                   
                   xlWorkSheet.Cells[Fila, 1] = Empleado;//INSERTAR INFORMACION EN CELDAS
                   xlWorkSheet.Cells[Fila, 2] = Total;//INSERTAR INFORMACION EN CELDAS
                    Fila++;
                }
            }

            // esto se puede omitir, pero lo dejo para ver cómo se podrían añadir/eliminar datos a posteriori
            // la primera fila en negrita, centrada y con fondo gris

            Excel.Range fila1 = (Excel.Range)xlWorkSheet.Rows[1];
            fila1.Select();
            fila1.EntireRow.Font.Bold = true;
            fila1.EntireRow.HorizontalAlignment = HorizontalAlignment.Center;
            fila1.EntireRow.Interior.Color = Color.Gray;
            fila1.EntireColumn.AutoFit();
            Excel.Range c1f1 = (Excel.Range)xlWorkSheet.Cells[1, 1];
            if (c1f1.Text == "")
            {
                Excel.Range columna1 = (Excel.Range)xlWorkSheet.Columns[1];
                columna1.Select();
                columna1.Delete();
            }

            // selecciono la primera celda de la primera columna
            Excel.Range c1 = (Excel.Range)xlWorkSheet.Cells[1, 1];
            c1.Select();

            grd.ClearSelection();
            grd.MultiSelect = false;
        }

        public void GuardarHistorial_Retardos(String Empleado, String Turno, DateTime Checada,String Total)
        {
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            //importar la libreria de widws form para abrir una ventana 
            SqlCommand cmd = new SqlCommand("GuardarHistorial_Retardos", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@Empleado", Empleado);
            cmd.Parameters.AddWithValue("@Turno", Turno);
            cmd.Parameters.AddWithValue("@Checada", Checada);
            cmd.Parameters.AddWithValue("@total", Total);
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Load(cmd.ExecuteReader());
            cn.Close();
        }

        public DataTable MostrarMesRetardos(DateTime Checada)
        {
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            //importar la libreria de widws form para abrir una ventana 
            SqlCommand cmd = new SqlCommand("MostrarMesRetardos", cn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@FechaChecada", Checada);
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Load(cmd.ExecuteReader());
            cn.Close();
            return dt;
        }






    }
}

