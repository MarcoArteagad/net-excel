using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace basicosExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //creamos un objeto de conexion llamado myConnection
                System.Data.OleDb.OleDbConnection MyConnection;
                //creamos y preparamos un objeto para ejecutar los comandos
                System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();
                //preparamos una variable para realizar el query
                string sql = null;

                //especificamos el provedor o conector y la ruta donde se encuentra nuestro archivo
                MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source= C:/Users/AldrichAAV/Documents/Libro6.xlsx ;Extended Properties=Excel 12.0");
                MyConnection.Open();
                myCommand.Connection = MyConnection;

                //realizamos el query donde meteremos los datos recuerda la hoja, las cabeceras(campos), y por ultimo los valores
                sql = "Insert into [Hoja1$] (FECHA,NOMBRE,REPORTE,TELEFONO,DOMICILIO,OBSERVACIONES) values ('6 noviembre 2017','jarek','4','7352090','ZACATEPEC','PUESTA A PUNTO')";
                myCommand.CommandText = sql;
                myCommand.ExecuteNonQuery();
                MyConnection.Close();
                MessageBox.Show("Datos Insertados Correctamente");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
