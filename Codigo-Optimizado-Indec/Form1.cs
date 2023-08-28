using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using Microsoft.Extensions.Options;
using Microsoft.Office.Interop.Excel;
using objExcel = Microsoft.Office.Interop.Excel;

namespace Codigo_Optimizado_Indec
{
    partial class Form1 : Form
    {

        public RutaPDF r = new RutaPDF(); //objeto de la clase RutaPDF

        public RutaTXT rt = new RutaTXT(); //objeto de la clase RutaTXT

        int contador = 0, cuadro = 0, EleccionObra;

        double ViejoCostoFinanciero, NuevoCostoFinanciero, PonderacionCostoFinanciero = 0.03, total, PonderacionTotal;

        string numero1 = "", numero2 = "";

        public ObraSieteItem ObraSieteItem = new ObraSieteItem();

        public ObraTablestaca ObraTablestaca = new ObraTablestaca();

        public Form1()
        {
            InitializeComponent();
            radioButtonDesaguesPluviales.Checked = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void buttonRutaPDF_Click(object sender, EventArgs e) //boton para buscar la ruta del pdf
        {
            r.ObtenerRuta(); //funcion para obtener la ruta
            textBoxRuta.Text = r.RutaArchivo; //guardamos la direccion de la ruta en el textbox

            if (!string.IsNullOrWhiteSpace(textBoxRuta.Text))
            {
                buttonPrimeraPagina.Enabled = true;
            }
            else
            {
                MessageBox.Show("No selecciono ningun archivo!!!!");
            }
        }

        public void CuadroGuardar(string CuadroGuardar)
        {
            rt.GuardarArchivoTXT(); //funcion para guardar el archivo de texto
            buttonContinuar.Enabled = true;
            buttontxt.Text = CuadroGuardar;
            buttontxt.Enabled = false;
        }

        private void buttontxt_Click(object sender, EventArgs e)
        {

            switch (contador)
            {
                case 0:
                    CuadroGuardar("Cuadro 5 guardar txt");
                    break;
                case 1:
                    CuadroGuardar("Cuadro 4 guardar txt");
                    break;
                case 2:
                    CuadroGuardar("Cuadro 3 guardar txt");
                    break;
                case 3:
                    CuadroGuardar("Cuadro 3 guardar txt");
                break;
            }
        }

        public void FuncionBotonContinuar() //funcion para obtener el contenido del pdf y escribirlo en archivos txt
        {
            var pdfDocument = new PdfDocument(new PdfReader(textBoxRuta.Text));
            var strategy = new LocationTextExtractionStrategy();
            r.Text = string.Empty;
            StreamWriter file = new StreamWriter(rt.Archivo, true);
            for (int i = 1; i <= pdfDocument.GetNumberOfPages(); i++) //for para obtener la cantidad de paginas
            {

                if (r.PrimeraPagina == i && r.UltimaPagina >= i) //if para obtener las paginas especificadas
                {
                    var page = pdfDocument.GetPage(r.PrimeraPagina++); //obtiene el numero de pagina dentro del pdf
                    r.Text = PdfTextExtractor.GetTextFromPage(page); //obtiene el texto dentro del pdf
                    file.Write(r.Text); //escribe las lineas de codigo dentro de los archivos de texto
                    Debug.WriteLine(r.Text);

                }

            }

            file.Close();
            file.Dispose();
        }

        public void CuadroCrearTxt(string CuadroCrear, bool VerdaderoFalso)
        {
            FuncionBotonContinuar(); //funcion para obtener el contenido del pdf y escribirlo en archivos txt
            buttonContinuar.Text = CuadroCrear;
            buttonContinuar.Enabled = false;
            contador++;
            buttonPrimeraPagina.Enabled = VerdaderoFalso;
        }

        private void buttonContinuar_Click(object sender, EventArgs e)
        {

            switch (contador)
            {
                case 0:
                    CuadroCrearTxt("Cuadro 5 crear txt", true);
                    break;
                case 1:
                    CuadroCrearTxt("Cuadro 4 crear txt", true);
                    break;
                case 2:
                    CuadroCrearTxt("Cuadro 3 crear txt", true);
                    break;
                case 3:
                    CuadroCrearTxt("Cuadro 3 crear txt", false);
                break;
            }
        }

        public void CuadroInicio(string CuadroInicio)
        {
            r.PrimeraPagina = int.Parse(maskedTextBoxPrimeraPagina.Text); //le pedimos al usuario la pagina donde inicia el cuadro
            buttonUltimaPagina.Enabled = true;
            buttonPrimeraPagina.Text = CuadroInicio;
            buttonPrimeraPagina.Enabled = false;
        }

        private void buttonPrimeraPagina_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(maskedTextBoxPrimeraPagina.Text))
            {
                switch (contador)
                {
                    case 0:
                        CuadroInicio("Cuadro 5 Inicio");
                        break;
                    case 1:
                        CuadroInicio("Cuadro 4 Inicio");
                        break;
                    case 2:
                        CuadroInicio("Cuadro 3 Inicio");
                        break;
                    case 3:
                        CuadroInicio("Cuadro 3 Inicio");
                        break;
                }
            }
            else
            {
                MessageBox.Show("Debe ingresar un valor!!!!");
            }
            
        }

        public void CuadroFin(string CuadroFin)
        {
            r.UltimaPagina = int.Parse(maskedTextBoxUltimaPagina.Text); //le pedimos al usuario la pagina donde finaliza el cuadro
            buttontxt.Enabled = true;
            buttonUltimaPagina.Text = CuadroFin;
            buttonUltimaPagina.Enabled = false;
        }

        private void buttonUltimaPagina_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(maskedTextBoxUltimaPagina.Text))
            {
                switch (contador)
                {
                    case 0:
                        CuadroFin("Cuadro 5 Fin");
                        break;
                    case 1:
                        CuadroFin("Cuadro 4 Fin");
                        break;
                    case 2:
                        CuadroFin("Cuadro 3 Fin");
                        break;
                    case 3:
                        CuadroFin("Cuadro 3 Fin");
                        break;
                }
            }
            else
            {
                MessageBox.Show("Debe ingresar un valor!!!!");
            }
        }

        string[] trozos;

        public void AgregarConCostoFinanciero(double ponderacion, int UltimaPosicion, int ViejaPosicion,EItems item, EItems TituloCostoFinanciero)
        {
            double variacion = double.Parse(trozos[UltimaPosicion]) / double.Parse(trozos[ViejaPosicion]);
            
            double IndiceVariacionResultante;
            IndiceVariacionResultante = ponderacion * variacion;
            dataGridView1.Rows.Add(item, ponderacion, trozos[ViejaPosicion], trozos[UltimaPosicion], variacion, IndiceVariacionResultante);

            double variacionFinanciera = NuevoCostoFinanciero / ViejoCostoFinanciero;
            double IndiceVariacionResultanteFinanciera = ponderacion * variacionFinanciera;
            dataGridView1.Rows.Add(TituloCostoFinanciero, PonderacionCostoFinanciero, ViejoCostoFinanciero, NuevoCostoFinanciero, variacionFinanciera, IndiceVariacionResultanteFinanciera);
            total = total + IndiceVariacionResultanteFinanciera;

            total = total + IndiceVariacionResultante;
            PonderacionTotal = PonderacionTotal + PonderacionCostoFinanciero;
            PonderacionTotal = PonderacionTotal + ponderacion;
        }

        public void RestoDeLosCuadros(double ponderacion, int UltimaPosicion, int ViejaPosicion, EItems item)
        {
            double variacion = double.Parse(trozos[UltimaPosicion]) / double.Parse(trozos[ViejaPosicion]);
            double IndiceVariacionResultante;
            IndiceVariacionResultante = ponderacion * variacion;
            dataGridView1.Rows.Add(item, ponderacion, trozos[ViejaPosicion], trozos[UltimaPosicion], variacion, IndiceVariacionResultante);
            total = total + IndiceVariacionResultante;
            PonderacionTotal = PonderacionTotal + ponderacion;
        }

        private void Parsear() //funcion para parsear los archivos de texto
        {
            trozos = rt.Linea.Split(' '); //asignamos que el separador es el espacio vacio
            trozos = trozos.ToList().Where(x => !string.IsNullOrEmpty(x)).ToArray(); //esto sirve para indicar que todo espacio vacio extra no nos moleste
            int i = 0;
            Debug.WriteLine(rt.Linea); //aca esbrico en el debug cada linea del archivo de texto
            dataGridView1.AllowUserToAddRows = false;
            
            while (i < trozos.Length)
            {
                Debug.WriteLine("[" + trozos[i] + "]"); //aca escribo en el debug como se ve parseado mostrando las separaciones con corchetes
                i++;
            }

            switch (cuadro)
            {
                case 1: //cuadro 1
                    switch (rt.NumeroLinea)
                    {
                        case 13:
                            dataGridView1.ColumnCount = 6; //asigno el numero de columnas
                            dataGridView1.Columns[0].HeaderText = "Insumos"; //agrego titulo a la columna 0
                            dataGridView1.Columns[1].HeaderText = "Ponderacion"; //agrego titulo a la columna 1
                            dataGridView1.Columns[2].HeaderText = trozos[12]; //aca se agrega en la columna 2 el mes anterior
                            dataGridView1.Columns[3].HeaderText = trozos[13]; //aca se agrega en la columna 3 el mes actual
                            dataGridView1.Columns[4].HeaderText = "Variacion"; //agrego titulo a la columna 4
                            dataGridView1.Columns[5].HeaderText = "Indice de variacion resultante"; //agrego titulo a la columna 5
                            break;
                        case 27:
                            if (EleccionObra == 1 || EleccionObra == 3)
                            {
                                AgregarConCostoFinanciero(0.09, 19, 18, EItems.Asfaltos_Combustibles_Lubricantes, EItems.Costo_Financiero);
                            }
                            else if (EleccionObra == 2 || EleccionObra == 5)
                            {
                                AgregarConCostoFinanciero(0.34, 19, 18, EItems.Asfaltos_Combustibles_Lubricantes, EItems.Costo_Financiero);
                            }
                            else if (EleccionObra == 4)
                            {
                                AgregarConCostoFinanciero(0.10, 19, 18, EItems.Asfaltos_Combustibles_Lubricantes, EItems.Costo_Financiero);
                            }
                            break;
                        case 40:
                            if (EleccionObra == 1 || EleccionObra == 3)
                            {
                                RestoDeLosCuadros(0.15, 16, 15, EItems.Equipo);
                            }
                            else if (EleccionObra == 2 || EleccionObra == 5)
                            {
                                RestoDeLosCuadros(0.35, 16, 15, EItems.Equipo);
                            }
                            else if (EleccionObra == 4)
                            {
                                RestoDeLosCuadros(0.10, 16, 15, EItems.Equipo);
                            }
                        break;
                    }

                    break;
                case 2: //cuadro 5
                    switch (rt.NumeroLinea)
                    {
                        case 16:
                            if (trozos[0] == "a)") //pregunto si en la linea 16 existe el indice a)
                            {
                                if (EleccionObra == 1)
                                {
                                    RestoDeLosCuadros(0.24, 16, 15, EItems.Mano_de_Obra);
                                }
                                else if (EleccionObra == 2 || EleccionObra == 4 || EleccionObra == 5)
                                {
                                    RestoDeLosCuadros(0.20, 16, 15, EItems.Mano_de_Obra);
                                }
                                else if (EleccionObra == 3)
                                {
                                    RestoDeLosCuadros(0.30, 16, 15, EItems.Mano_de_Obra);
                                }
                            }
                            break;
                        case 17:
                            if (trozos[0] == "a)") //pregunto si en la linea 17 existe el indice a)
                            {
                                if (EleccionObra == 1)
                                {
                                    RestoDeLosCuadros(0.24, 16, 15, EItems.Mano_de_Obra);
                                }
                                else if (EleccionObra == 2 || EleccionObra == 4 || EleccionObra == 5)
                                {
                                    RestoDeLosCuadros(0.20, 16, 15, EItems.Mano_de_Obra);
                                }
                                else if (EleccionObra == 3)
                                {
                                    RestoDeLosCuadros(0.30, 16, 15, EItems.Mano_de_Obra);
                                }
                            }
                            break;
                        case 39:
                            if (trozos[0] == "p)") //pregunto si en la linea 39 existe el indice p)
                            {
                                if (EleccionObra == 1 || EleccionObra == 2 || EleccionObra == 3)
                                {
                                    RestoDeLosCuadros(0.08, 16, 15, EItems.Gasto_General);
                                }
                                else if (EleccionObra == 4 || EleccionObra == 5)
                                {
                                    RestoDeLosCuadros(0.15, 16, 15, EItems.Gasto_General);
                                }
                            }
                            break;
                        case 40:
                            if (trozos[0] == "p)") //pregunto si en la linea 40 existe el indice p)
                            {
                                if (EleccionObra == 1 || EleccionObra == 2 || EleccionObra == 3)
                                {
                                    RestoDeLosCuadros(0.08, 16, 15, EItems.Gasto_General);
                                }
                                else if (EleccionObra == 4 || EleccionObra == 5)
                                {
                                    RestoDeLosCuadros(0.15, 16, 15, EItems.Gasto_General);
                                }
                            }
                            break;
                        case 45:
                            if (trozos[0] == "s)") //pregunto si en la linea 39 existe el indice s)
                            {
                                if (EleccionObra == 1)
                                {
                                    RestoDeLosCuadros(0.30, 16, 15, EItems.Hormigon);
                                }
                                else if (EleccionObra == 3)
                                {
                                    RestoDeLosCuadros(0.22, 16, 15, EItems.Hormigon);
                                }
                                else if (EleccionObra == 4)
                                {
                                    RestoDeLosCuadros(0.12, 16, 15, EItems.Hormigon);
                                }
                            }
                            break;
                        case 46:
                            if (trozos[0] == "s)") //pregunto si en la linea 39 existe el indice s)
                            {
                                if (EleccionObra == 1)
                                {
                                    RestoDeLosCuadros(0.30, 16, 15, EItems.Hormigon);
                                }
                                else if (EleccionObra == 3)
                                {
                                    RestoDeLosCuadros(0.22, 16, 15, EItems.Hormigon);
                                }
                                else if (EleccionObra == 4)
                                {
                                    RestoDeLosCuadros(0.12, 16, 15, EItems.Hormigon);
                                }
                            }
                            break;
                    }
                    break;
                case 3: //cuadro 4
                    switch (rt.NumeroLinea)
                    {
                        case 20:
                            numero1 = trozos[2].ToString();
                            break;
                        case 21:
                            numero2 = trozos[2].ToString();

                            if (EleccionObra == 1)
                            {
                                double variacion = double.Parse(numero2) / double.Parse(numero1);
                                double ponderacion = 0.11;
                                double IndiceVariacionResultante;
                                IndiceVariacionResultante = ponderacion * variacion;
                                dataGridView1.Rows.Add(EItems.Acero, ponderacion, numero1, numero2, variacion, IndiceVariacionResultante);
                                total = total + IndiceVariacionResultante;
                                PonderacionTotal = PonderacionTotal + ponderacion;
                            }
                            else if (EleccionObra == 3)
                            {
                                double variacion = double.Parse(numero2) / double.Parse(numero1);
                                double ponderacion = 0.13;
                                double IndiceVariacionResultante;
                                IndiceVariacionResultante = ponderacion * variacion;
                                dataGridView1.Rows.Add(EItems.Acero, ponderacion, numero1, numero2, variacion, IndiceVariacionResultante);
                                total = total + IndiceVariacionResultante;
                                PonderacionTotal = PonderacionTotal + ponderacion;
                            }
                        break;
                    }
                    break;
                case 4: //cuadro 4
                    if (rt.NumeroLinea == 24)
                    {
                        if (EleccionObra == 4)
                        {
                            RestoDeLosCuadros(0.30, 19, 18, EItems.Tablestaca);
                        }
                    }
                break;
            }
        }

        StreamReader LeerLineas;

        public void RecorrerLinea(int NumeroDeLinea)
        {

            while (!LeerLineas.EndOfStream) //while que recorre el cuadro por linea hasta el final del archivo
            {
                rt.Linea = LeerLineas.ReadLine();

                if (++rt.NumeroLinea == NumeroDeLinea) //if para obtener la linea especifica dentro del archivo de texto
                {
                    Parsear(); //funcion para parsear los archivos de texto
                    break;

                }
            }
        }

        private void buttonPruebas_Click(object sender, EventArgs e)
        {
            
            radioButtonDesaguesPluviales.Enabled = false;
            radioButtonExcavacionCanal.Enabled = false;
            radioButtonPresas.Enabled = false;
            radioButtonDefensaCostera.Enabled = false;
            radioButtonDefensaPoblacion.Enabled = false;

            LeerLineas = File.OpenText(textBoxRutaTXT.Text);

            switch (cuadro)
            {
                case 1: //cuadro 1
                    RecorrerLinea(13);
                    RecorrerLinea(27);
                    RecorrerLinea(40);
                    rt.NumeroLinea = 0;
                    break;
                case 2: //cuadro 5
                    RecorrerLinea(16);
                    RecorrerLinea(17);
                    RecorrerLinea(39);
                    RecorrerLinea(40);
                    RecorrerLinea(45);
                    RecorrerLinea(46);
                    rt.NumeroLinea = 0;
                    break; 
                case 3: //cuadro 4
                    RecorrerLinea(20);
                    RecorrerLinea(21);
                    rt.NumeroLinea = 0;
                    break; 
                case 4: //cuadro 3
                    RecorrerLinea(24);
                    break;
            }
            buttonPruebas.Enabled = false;
            buttonRutaTXT.Enabled = true;
        }

        private void buttonRutaTXT_Click(object sender, EventArgs e) //boton para obtener la ruta del archivo de texto que queremos analizar
        {
            rt.ObtenerRutaTXT(); //funcion para obtener la ruta donde se encuentran guardados los archivos de texto
            textBoxRutaTXT.Text = rt.RutaArchivoTXT; //guardamos la direccion de la ruta en el textbox

            if (!string.IsNullOrWhiteSpace(textBoxRutaTXT.Text))
            {
                

                cuadro++;
                buttonRutaTXT.Enabled = false;
                buttonPruebas.Enabled = true;
            }
            else
            {
                MessageBox.Show("No selecciono ningun archivo!!!!");
            } 
        }

        private void buttonExportarExcel_Click(object sender, EventArgs e) //boton para exportar los datos del datagridview al excel
        {
            objExcel.Application application = new objExcel.Application();
            Workbook objLibro = application.Workbooks.Add(XlSheetType.xlWorksheet);
            Worksheet objHoja = (Worksheet)application.ActiveSheet;

            application.Visible = true;

            foreach (DataGridViewColumn columna in dataGridView1.Columns)
            {
                objHoja.Cells[1, columna.Index + 1] = columna.HeaderText;
                foreach(DataGridViewRow fila in dataGridView1.Rows)
                {
                    objHoja.Cells[fila.Index + 2, columna.Index + 1] = fila.Cells[columna.Index].Value;
                }
            }
        }

        private void radioButtonDesaguesPluviales_CheckedChanged(object sender, EventArgs e)
        {
            EleccionObra = 1;
        }

        private void radioButtonExcavacionCanal_CheckedChanged(object sender, EventArgs e)
        {
            EleccionObra = 2;
        }

        private void buttonTotal_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add("Total", PonderacionTotal, "", "", "", total);
            buttonExportarExcel.Enabled = true;
            buttonRutaTXT.Enabled = false;
        }

        private void radioButtonPresas_CheckedChanged(object sender, EventArgs e)
        {
            EleccionObra = 3;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(button2.Text))
            {
                NuevoCostoFinanciero = double.Parse(maskedNuevoCostoFinanciero.Text);
                button2.Enabled = false;
            }
            else
            {
                MessageBox.Show("Debe ingresar un valor!!!!");
            }
        }

        private void radioButtonDefensaCostera_CheckedChanged(object sender, EventArgs e)
        {
            EleccionObra = 4;
        }

        private void radioButtonDefensaPoblacion_CheckedChanged(object sender, EventArgs e)
        {
            EleccionObra = 5;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrWhiteSpace(button1.Text))
            {
                ViejoCostoFinanciero = double.Parse(maskedViejoCostoFinanciero.Text);
                button1.Enabled = false;
            }
            else
            {
                MessageBox.Show("Debe ingresar un valor!!!!");
            }
        }
    }
}
