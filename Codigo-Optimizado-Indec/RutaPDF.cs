using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Codigo_Optimizado_Indec
{
    public class RutaPDF
    {

        private string rutaArchivo;

        public string RutaArchivo
        {
            get { return rutaArchivo; }
            set { rutaArchivo = value; }
        }

        private int primeraPagina;

        public int PrimeraPagina
        {
            get { return primeraPagina; }
            set { primeraPagina = value; }
        }

        private int ultimaPagina;

        public int UltimaPagina
        {
            get { return ultimaPagina; }
            set { ultimaPagina = value; }
        }


        private string text;

        public string Text
        {
            get { return text; }
            set { text = value; }
        }

        public string ObtenerRuta() //esta funcion se utiliza para obtener la rtua donde se encuentra el pdf
        {
            OpenFileDialog OpenFileDialog = new OpenFileDialog();

            if (OpenFileDialog.ShowDialog() == DialogResult.OK)
            {
                rutaArchivo = OpenFileDialog.FileName;
            }

            return rutaArchivo;
        }

    }
}