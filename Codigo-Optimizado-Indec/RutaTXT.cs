using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace Codigo_Optimizado_Indec
{
    public class RutaTXT
    {

        private string archivo;

        public string Archivo
        {
            get { return archivo; }
            set { archivo = value; }
        }

        private string rutaArchivoTXT;

        public string RutaArchivoTXT
        {
            get { return rutaArchivoTXT; }
            set { rutaArchivoTXT = value; }
        }

        private string linea = null;

        public string Linea
        {
            get { return linea; }
            set { linea = value; }
        }

        private int numeroLinea;

        public int NumeroLinea
        {
            get { return numeroLinea; }
            set { numeroLinea = value; }
        }

        public string GuardarArchivoTXT() //esta funcion sirve para guardar el archivo de texto
        {

            SaveFileDialog saveFileDialog = new SaveFileDialog();


            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (File.Exists(saveFileDialog.FileName)) //aca se pregunta si el archivo de texto existe
                {

                }

                archivo = saveFileDialog.FileName;
            }
            return archivo;
        }

        public string ObtenerRutaTXT() //funcion para obtener la ruta donde se encuentran guardados los archivos de texto
        {
            OpenFileDialog OpenFileDialog = new OpenFileDialog();

            if (OpenFileDialog.ShowDialog() == DialogResult.OK)
            {
                rutaArchivoTXT = OpenFileDialog.FileName;
            }

            return rutaArchivoTXT;
        }

    }
}