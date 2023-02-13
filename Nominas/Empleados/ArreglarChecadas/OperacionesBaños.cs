using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Collections;
using System.Drawing;

using System.Diagnostics;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Point = System.Drawing.Point;
using DataTable = System.Data.DataTable;
using System.Data.SqlClient;//libreria sql
using System.Data;//libreria sql 
using System.Data.Sql;//libreria sql
using System.Windows.Forms;//libreria para usar los messagebox
using System.Data.OleDb;
using System.Globalization;

namespace Nominas
{
    class OperacionesBaños
    {
        CadenaConexion var = new CadenaConexion();
        SqlConnection cn; // para la conexion a SQL
        SqlCommand cmd;  // para las instrucciones SQL
        SqlDataReader dr; //

        public void ImnportarBaños(String Tarjeta, String Nombre,String Checada, String Tipo)
        {
            try
            {
                DateTime Fecha = Convert.ToDateTime(Checada);
                if (Tipo.EndsWith("Entrada"))
                {
                    Tipo = "Entrada";
                }
                else if (Tipo.EndsWith("Salida"))
                {
                    Tipo = "Salida";
                }

                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                SqlCommand cmd = new SqlCommand("CargarBaños", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Tarjeta", Tarjeta);
                cmd.Parameters.AddWithValue("@Nombre", Nombre);
                cmd.Parameters.AddWithValue("@Fecha", Fecha);
                cmd.Parameters.AddWithValue("@Tipo", Tipo);
                cmd.ExecuteReader();
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error falta: " + ex.ToString());
            }
        }

        public DataTable ConsultarBaños()
        {
            DataTable dt = new DataTable();
            try
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                SqlCommand cmd = new SqlCommand("ConsultarBaños", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                dt.Load(cmd.ExecuteReader());
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error falta: " + ex.ToString());
            }
            return dt;
        }

        public DataTable ConsultarHorasBaños()
        {
            DataTable dt = new DataTable();
            try
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                SqlCommand cmd = new SqlCommand("ConsultarHorassBaños", cn);
                cmd.CommandType = CommandType.StoredProcedure;
                dt.Load(cmd.ExecuteReader());
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error falta: " + ex.ToString());
            }
            return dt;
        }

        public void DuplicadosBaños()
        {
            ArrayList Lista = new ArrayList();
            DataTable dt = new DataTable();
            DataTable dtSemana = new DataTable();
            dt = this.ListaUnicaBaño();

            MessageBox.Show("Contador de lista unica: " + dt.Rows.Count);  
           
            for (int T = 0; T < dt.Rows.Count; T++)
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                cmd = new SqlCommand("SELECT * FROM Baños WHERE Tarjeta = '" + dt.Rows[T]["Tarjeta"].ToString() + "' order by Checada", cn);
                dtSemana.Load(cmd.ExecuteReader());
                cn.Close();
                
                int Limite = dtSemana.Rows.Count;
                for (int x = 0; x < dtSemana.Rows.Count; x++)
                {
                    if (x < (Limite - 1))
                    {
                        String id = dtSemana.Rows[x]["id"].ToString();
                        DateTime Checada = Convert.ToDateTime(dtSemana.Rows[x]["Checada"].ToString());
                        String Tipo = dtSemana.Rows[x]["Tipo"].ToString();

                        String id2 = dtSemana.Rows[x + 1]["id"].ToString();
                        DateTime Checada2 = Convert.ToDateTime(dtSemana.Rows[x + 1]["Checada"].ToString());
                        String Tipo2 = dtSemana.Rows[x + 1]["Tipo"].ToString();

                        if (x == 0 && Tipo.Equals("Salida") )
                        {
                            Lista.Add(id);
                        }
                        else if (Checada == Checada2 && Tipo.Equals(Checada2))
                        {
                            Lista.Add(id);
                        }
                        // MessageBox.Show("ID: " + id + " Segundo: " + id2);
                        else if (Tipo2.Equals(Tipo) && Checada < Checada2)
                        {
                            Lista.Add(id2);
                            Lista.Add(id);
                        }
                        else if ((Tipo2.Equals(Tipo) && Checada2 < Checada.AddMinutes(+1)))
                        {
                           // MessageBox.Show(x.ToString() + " Duplicadas de : " + dt.Rows[T]["Tarjeta"].ToString()  + "   " + Checada.ToString() + " y " + Checada2.ToString() + " Eliminar : " + id2);
                            Lista.Add(id2);
                        }
                    }
                }//semana empleado

                for (int f=0; f < Lista.Count; f++)
                {
                    cn = new SqlConnection(var.GetConexionDB());
                    cn.Open();
                    cmd = new SqlCommand("DELETE Baños WHERE id = '" + Lista[f].ToString() + "'", cn);
                    dtSemana.Load(cmd.ExecuteReader());
                    cn.Close();
                }
                dtSemana.Clear();
            }//lista Empleados
        }//Fin de funcion

        public void CalcularBaños()
        {
            ArrayList Lista = new ArrayList();
            DataTable dt = new DataTable();
            DataTable dtSemana = new DataTable();

            ArrayList Lunes = new ArrayList();
            ArrayList Martes = new ArrayList();
            ArrayList Miercoles = new ArrayList();
            ArrayList Jueves = new ArrayList();
            ArrayList Viernes = new ArrayList();
            ArrayList Sabado = new ArrayList();

            Double TotalHoras = 0;
            var horas = 0.0;
            ////////////////// OBTENER LISTA SIN DUPLICADOS ////////////////////////
            try
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                cmd = new SqlCommand("TRUNCATE TABLE CalculoBaños", cn);
                cmd.ExecuteReader();
                cn.Close();

                cn.Open();
                cmd = new SqlCommand("SELECT DISTINCT Tarjeta, Nombre FROM Baños ORDER BY Tarjeta ASC", cn);
                dt.Load(cmd.ExecuteReader());
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("No se conecto" + ex.ToString());
            }
            
            ////////////////// REGISTRAR EN TABLA DE CALCULO ////////////////////////
            
            for (int x = 0; x < dt.Rows.Count; x++)
            {
                cn = new SqlConnection(var.GetConexionDB());
                cn.Open();
                cmd = new SqlCommand("INSERT INTO CalculoBaños(Tarjeta,Nombre) VALUES ('" + dt.Rows[x]["Tarjeta"].ToString() + "','" + dt.Rows[x]["Nombre"].ToString() + "')", cn);
                cmd.ExecuteReader();
                cn.Close();
            }

            ////////////////// OBTENER LISTA DE OPERADORES ////////////////////////
            try
            {
                dt = this.ListaUnicaBaño();
                for (int T = 0; T < dt.Rows.Count; T++)
                {
                    cn = new SqlConnection(var.GetConexionDB());
                    cn.Open();
                    cmd = new SqlCommand("SELECT * FROM Baños WHERE Tarjeta = '" + dt.Rows[T]["Tarjeta"].ToString() + "'  order by Checada", cn);
                    dtSemana.Load(cmd.ExecuteReader());
                    cn.Close();
                    TotalHoras = 0;

                    ////////////////// OBTENER SEMANA DE OPERADORES ////////////////////////
                    //MessageBox.Show(dtSemana.Rows.Count.ToString() + " De: " + dt.Rows[T]["Tarjeta"].ToString());
                    for (int j = 0; j < dtSemana.Rows.Count; j+=2)
                    {
                        if (j<dtSemana.Rows.Count)
                        {
                            if (j < (dtSemana.Rows.Count) && dtSemana.Rows[j]["Tipo"].ToString().Equals("Entrada"))
                            {
                                horas = (Convert.ToDateTime(dtSemana.Rows[j + 1]["Checada"].ToString()) - Convert.ToDateTime(dtSemana.Rows[j]["Checada"].ToString())).TotalMinutes;

                                TotalHoras = TotalHoras + horas;
                               // MessageBox.Show("Tarjeta: " + dtSemana.Rows[j]["Tarjeta"].ToString() + " " + dtSemana.Rows[j + 1]["Checada"].ToString() + " - " + dtSemana.Rows[j]["Checada"].ToString() + " Total: " + TotalHoras.ToString());
                            }
                        }
                    }//Fin de for de seleccion de dia
                    //MessageBox.Show("Horas: " + TotalHoras.ToString() + "Tarjeta " + dt.Rows[T]["Tarjeta"].ToString());
                    dtSemana.Clear();

                    cn.Open();
                    cmd = new SqlCommand("UPDATE CalculoBaños SET Total = '" + TotalHoras.ToString() + "' WHERE Tarjeta = '" + dt.Rows[T]["Tarjeta"].ToString() + "'", cn);
                    cmd.ExecuteReader();
                    cn.Close();
                    TotalHoras = 0;
                    horas = 0;
                }//For de Lista por trabajador

            }catch (Exception ex)
            {
                MessageBox.Show("No se conecto" + ex.ToString());
            }
                /*
                if (x < (LimiteSemana - 1))
                {
                    String id = dtSemana.Rows[x]["id"].ToString();
                    DateTime Checada = Convert.ToDateTime(dtSemana.Rows[x]["Checada"].ToString());
                    String Tipo = dtSemana.Rows[x]["Tipo"].ToString();

                    String id2 = dtSemana.Rows[x + 1]["id"].ToString();
                    DateTime Checada2 = Convert.ToDateTime(dtSemana.Rows[x + 1]["Checada"].ToString());
                    String Tipo2 = dtSemana.Rows[x + 1]["Tipo"].ToString();
                }*/
        }




        public DataTable ListaUnicaBaño()
        {
            DataTable dt = new DataTable();
            cn = new SqlConnection(var.GetConexionDB());
            cn.Open();
            cmd = new SqlCommand("SELECT DISTINCT Tarjeta, Nombre FROM Baños ORDER BY Tarjeta ASC", cn);
            dt.Load(cmd.ExecuteReader());
            cn.Close();

            return dt;
        }


        public void ExportarExcelLineas(DataGridView grd)
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
        }//Fin de ExportarExcel
    }
}
