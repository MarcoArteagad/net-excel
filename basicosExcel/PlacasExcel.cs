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
    public partial class PlacasExcel : Form
    {
        
        
        public PlacasExcel()
        {
            InitializeComponent();
            txtPlacas.Text = GenerarCodigo();

        }

        private string GenerarCodigo()
        {
            Random obj = new Random();
            string sCadena = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890";
            int longitud = sCadena.Length;
            char cletra;
            int nlongitud = 5;
            string sNuevacadena = string.Empty;
            //string posibles[];

            for (int i = 0; i < nlongitud; i++)
            {
                cletra = sCadena[obj.Next(longitud)];
                sNuevacadena += cletra;
            }
            return sNuevacadena;

        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            try
            {
                //creamos un objeto de conexion llamado myConnection
                System.Data.OleDb.OleDbConnection MyConnection;
                //creamos y preparamos un objeto para ejecutar los comandos
                System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();
                //preparamos una variable para realizar el query
                string sql = null;
                string serie=txtSerie.Text;
                string propietario=txtPropietario.Text;
                string vehiculo=txtVehiculo.Text;
                string modelo=comboBox1.Text;
                string puertas=comboBox2.Text;
                string año=txtAño.Text;
                string estado=txtEstado.Text;
                string placas=txtPlacas.Text;
                
                

                //especificamos el provedor o conector y la ruta donde se encuentra nuestro archivo
                MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source= C:/Users/AldrichAAV/Documents/Placas.xlsx ;Extended Properties=Excel 12.0");
                MyConnection.Open();
                myCommand.Connection = MyConnection;

                //realizamos el query donde meteremos los datos recuerda la hoja, las cabeceras(campos), y por ultimo los valores
                sql = "Insert into [Hoja1$] (Serie,Propietario,Vehiculo,Modelo,Puertas,Año,Estado,Placas) values ('"+serie+"','" + propietario + "','" + vehiculo + "','" + modelo + "','" + puertas + "','" + año + "','" + estado + "','" + placas +"')";
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
