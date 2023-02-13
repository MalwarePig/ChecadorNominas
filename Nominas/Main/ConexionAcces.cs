using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Nominas
{
    class ConexionAcces
    {
        public String GetConexionAcces()
        {
            String conec = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= C:\\Users\\Public\\Documents\\IBIX\\Datos\\Checador.mdb; Jet OLEDB:Database Password= kratos;";

            return conec;
        }
    }
}
