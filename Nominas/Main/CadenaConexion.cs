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
namespace Nominas
{
    class CadenaConexion
    {
        private String ConexionDB;
        SqlConnection cn; // para la conexion a SQL
        SqlCommand cmd;  // para las instrucciones SQL
        SqlDataReader dr; //

        public String GetConexionDB()
        {
           // this.ConexionDB = "Data Source = MAQUINA02\\SQLEXPRESS; Initial Catalog = LocalDB; User ID = SoporteDB; Password = S1702861";
            //this.ConexionDB = "Data Source=SOPORTE\\SQLEXPRESS01;Initial Catalog=LocalDB;Integrated Security=True";
            //this.ConexionDB = "Data Source = SERVIDORNOMINA\\SQLEXPRESS; Initial Catalog = LocalDB; User ID = SoporteDB; Password = S1702861";//Prime
            this.ConexionDB = "Data Source= PCMASTERRACE\\SQLEXPRESS;Initial Catalog=LocalDB;Integrated Security=True"; //Pruebas
            //this.ConexionDB = "Data Source=SERVIDOR;Initial Catalog=GemakDB;Persist Security Info=True;User ID=Soporte;Password=enigma1702861";
            return this.ConexionDB;
        }
    }
}
