namespace Codigo_Optimizado_Indec
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.labelTitulo = new System.Windows.Forms.Label();
            this.labelRutaPdf = new System.Windows.Forms.Label();
            this.textBoxRuta = new System.Windows.Forms.TextBox();
            this.buttonRutaPDF = new System.Windows.Forms.Button();
            this.buttontxt = new System.Windows.Forms.Button();
            this.buttonContinuar = new System.Windows.Forms.Button();
            this.buttonPrimeraPagina = new System.Windows.Forms.Button();
            this.buttonUltimaPagina = new System.Windows.Forms.Button();
            this.maskedTextBoxPrimeraPagina = new System.Windows.Forms.MaskedTextBox();
            this.maskedTextBoxUltimaPagina = new System.Windows.Forms.MaskedTextBox();
            this.buttonPruebas = new System.Windows.Forms.Button();
            this.textBoxRutaTXT = new System.Windows.Forms.TextBox();
            this.labelRutaTXT = new System.Windows.Forms.Label();
            this.buttonRutaTXT = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.buttonExportarExcel = new System.Windows.Forms.Button();
            this.radioButtonDesaguesPluviales = new System.Windows.Forms.RadioButton();
            this.radioButtonExcavacionCanal = new System.Windows.Forms.RadioButton();
            this.radioButtonPresas = new System.Windows.Forms.RadioButton();
            this.radioButtonDefensaCostera = new System.Windows.Forms.RadioButton();
            this.radioButtonDefensaPoblacion = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.maskedViejoCostoFinanciero = new System.Windows.Forms.MaskedTextBox();
            this.maskedNuevoCostoFinanciero = new System.Windows.Forms.MaskedTextBox();
            this.buttonTotal = new System.Windows.Forms.Button();
            this.labelCredito = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // labelTitulo
            // 
            this.labelTitulo.AutoSize = true;
            this.labelTitulo.Font = new System.Drawing.Font("MV Boli", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelTitulo.Location = new System.Drawing.Point(325, 8);
            this.labelTitulo.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelTitulo.Name = "labelTitulo";
            this.labelTitulo.Size = new System.Drawing.Size(509, 29);
            this.labelTitulo.TabIndex = 4;
            this.labelTitulo.Text = "Calculo en la Variación de Obras Hidráulicas";
            this.labelTitulo.UseWaitCursor = true;
            // 
            // labelRutaPdf
            // 
            this.labelRutaPdf.AutoSize = true;
            this.labelRutaPdf.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelRutaPdf.Location = new System.Drawing.Point(9, 39);
            this.labelRutaPdf.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelRutaPdf.Name = "labelRutaPdf";
            this.labelRutaPdf.Size = new System.Drawing.Size(167, 20);
            this.labelRutaPdf.TabIndex = 6;
            this.labelRutaPdf.Text = "Ingrese la ruta del pdf:";
            // 
            // textBoxRuta
            // 
            this.textBoxRuta.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxRuta.Location = new System.Drawing.Point(198, 40);
            this.textBoxRuta.Margin = new System.Windows.Forms.Padding(2);
            this.textBoxRuta.Name = "textBoxRuta";
            this.textBoxRuta.ReadOnly = true;
            this.textBoxRuta.Size = new System.Drawing.Size(360, 23);
            this.textBoxRuta.TabIndex = 5;
            // 
            // buttonRutaPDF
            // 
            this.buttonRutaPDF.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonRutaPDF.Location = new System.Drawing.Point(9, 75);
            this.buttonRutaPDF.Margin = new System.Windows.Forms.Padding(2);
            this.buttonRutaPDF.Name = "buttonRutaPDF";
            this.buttonRutaPDF.Size = new System.Drawing.Size(548, 43);
            this.buttonRutaPDF.TabIndex = 7;
            this.buttonRutaPDF.Text = "Buscar PDF";
            this.buttonRutaPDF.UseVisualStyleBackColor = true;
            this.buttonRutaPDF.Click += new System.EventHandler(this.buttonRutaPDF_Click);
            // 
            // buttontxt
            // 
            this.buttontxt.Enabled = false;
            this.buttontxt.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttontxt.Location = new System.Drawing.Point(9, 290);
            this.buttontxt.Margin = new System.Windows.Forms.Padding(2);
            this.buttontxt.Name = "buttontxt";
            this.buttontxt.Size = new System.Drawing.Size(548, 43);
            this.buttontxt.TabIndex = 8;
            this.buttontxt.Text = "Cuadro 1 ruta txt";
            this.buttontxt.UseVisualStyleBackColor = true;
            this.buttontxt.Click += new System.EventHandler(this.buttontxt_Click);
            // 
            // buttonContinuar
            // 
            this.buttonContinuar.Enabled = false;
            this.buttonContinuar.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonContinuar.Location = new System.Drawing.Point(9, 338);
            this.buttonContinuar.Margin = new System.Windows.Forms.Padding(2);
            this.buttonContinuar.Name = "buttonContinuar";
            this.buttonContinuar.Size = new System.Drawing.Size(548, 43);
            this.buttonContinuar.TabIndex = 9;
            this.buttonContinuar.Text = "Cuadro 1 crear txt";
            this.buttonContinuar.UseVisualStyleBackColor = true;
            this.buttonContinuar.Click += new System.EventHandler(this.buttonContinuar_Click);
            // 
            // buttonPrimeraPagina
            // 
            this.buttonPrimeraPagina.Enabled = false;
            this.buttonPrimeraPagina.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonPrimeraPagina.Location = new System.Drawing.Point(9, 123);
            this.buttonPrimeraPagina.Margin = new System.Windows.Forms.Padding(2);
            this.buttonPrimeraPagina.Name = "buttonPrimeraPagina";
            this.buttonPrimeraPagina.Size = new System.Drawing.Size(197, 43);
            this.buttonPrimeraPagina.TabIndex = 10;
            this.buttonPrimeraPagina.Text = "Cuadro 1 Inicio";
            this.buttonPrimeraPagina.UseVisualStyleBackColor = true;
            this.buttonPrimeraPagina.Click += new System.EventHandler(this.buttonPrimeraPagina_Click);
            // 
            // buttonUltimaPagina
            // 
            this.buttonUltimaPagina.Enabled = false;
            this.buttonUltimaPagina.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonUltimaPagina.Location = new System.Drawing.Point(9, 171);
            this.buttonUltimaPagina.Margin = new System.Windows.Forms.Padding(2);
            this.buttonUltimaPagina.Name = "buttonUltimaPagina";
            this.buttonUltimaPagina.Size = new System.Drawing.Size(197, 43);
            this.buttonUltimaPagina.TabIndex = 11;
            this.buttonUltimaPagina.Text = "Cuadro 1 Fin";
            this.buttonUltimaPagina.UseVisualStyleBackColor = true;
            this.buttonUltimaPagina.Click += new System.EventHandler(this.buttonUltimaPagina_Click);
            // 
            // maskedTextBoxPrimeraPagina
            // 
            this.maskedTextBoxPrimeraPagina.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.maskedTextBoxPrimeraPagina.Location = new System.Drawing.Point(221, 135);
            this.maskedTextBoxPrimeraPagina.Margin = new System.Windows.Forms.Padding(2);
            this.maskedTextBoxPrimeraPagina.Mask = "99999";
            this.maskedTextBoxPrimeraPagina.Name = "maskedTextBoxPrimeraPagina";
            this.maskedTextBoxPrimeraPagina.Size = new System.Drawing.Size(130, 23);
            this.maskedTextBoxPrimeraPagina.TabIndex = 12;
            this.maskedTextBoxPrimeraPagina.ValidatingType = typeof(int);
            // 
            // maskedTextBoxUltimaPagina
            // 
            this.maskedTextBoxUltimaPagina.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.maskedTextBoxUltimaPagina.Location = new System.Drawing.Point(221, 183);
            this.maskedTextBoxUltimaPagina.Margin = new System.Windows.Forms.Padding(2);
            this.maskedTextBoxUltimaPagina.Mask = "99999";
            this.maskedTextBoxUltimaPagina.Name = "maskedTextBoxUltimaPagina";
            this.maskedTextBoxUltimaPagina.Size = new System.Drawing.Size(130, 23);
            this.maskedTextBoxUltimaPagina.TabIndex = 13;
            this.maskedTextBoxUltimaPagina.ValidatingType = typeof(int);
            // 
            // buttonPruebas
            // 
            this.buttonPruebas.Enabled = false;
            this.buttonPruebas.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonPruebas.Location = new System.Drawing.Point(572, 290);
            this.buttonPruebas.Margin = new System.Windows.Forms.Padding(2);
            this.buttonPruebas.Name = "buttonPruebas";
            this.buttonPruebas.Size = new System.Drawing.Size(568, 43);
            this.buttonPruebas.TabIndex = 14;
            this.buttonPruebas.Text = "Agregar Valores";
            this.buttonPruebas.UseVisualStyleBackColor = true;
            this.buttonPruebas.Click += new System.EventHandler(this.buttonPruebas_Click);
            // 
            // textBoxRutaTXT
            // 
            this.textBoxRutaTXT.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxRutaTXT.Location = new System.Drawing.Point(198, 384);
            this.textBoxRutaTXT.Margin = new System.Windows.Forms.Padding(2);
            this.textBoxRutaTXT.Name = "textBoxRutaTXT";
            this.textBoxRutaTXT.ReadOnly = true;
            this.textBoxRutaTXT.Size = new System.Drawing.Size(360, 23);
            this.textBoxRutaTXT.TabIndex = 15;
            // 
            // labelRutaTXT
            // 
            this.labelRutaTXT.AutoSize = true;
            this.labelRutaTXT.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelRutaTXT.Location = new System.Drawing.Point(22, 384);
            this.labelRutaTXT.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelRutaTXT.Name = "labelRutaTXT";
            this.labelRutaTXT.Size = new System.Drawing.Size(167, 20);
            this.labelRutaTXT.TabIndex = 16;
            this.labelRutaTXT.Text = "Ingrese la ruta del pdf:";
            // 
            // buttonRutaTXT
            // 
            this.buttonRutaTXT.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonRutaTXT.Location = new System.Drawing.Point(9, 410);
            this.buttonRutaTXT.Margin = new System.Windows.Forms.Padding(2);
            this.buttonRutaTXT.Name = "buttonRutaTXT";
            this.buttonRutaTXT.Size = new System.Drawing.Size(548, 43);
            this.buttonRutaTXT.TabIndex = 17;
            this.buttonRutaTXT.Text = "Buscar TXT";
            this.buttonRutaTXT.UseVisualStyleBackColor = true;
            this.buttonRutaTXT.Click += new System.EventHandler(this.buttonRutaTXT_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(572, 39);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(2);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(568, 241);
            this.dataGridView1.TabIndex = 18;
            // 
            // buttonExportarExcel
            // 
            this.buttonExportarExcel.Enabled = false;
            this.buttonExportarExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonExportarExcel.Location = new System.Drawing.Point(572, 338);
            this.buttonExportarExcel.Margin = new System.Windows.Forms.Padding(2);
            this.buttonExportarExcel.Name = "buttonExportarExcel";
            this.buttonExportarExcel.Size = new System.Drawing.Size(568, 43);
            this.buttonExportarExcel.TabIndex = 19;
            this.buttonExportarExcel.Text = "Exportar Excel";
            this.buttonExportarExcel.UseVisualStyleBackColor = true;
            this.buttonExportarExcel.Click += new System.EventHandler(this.buttonExportarExcel_Click);
            // 
            // radioButtonDesaguesPluviales
            // 
            this.radioButtonDesaguesPluviales.AutoSize = true;
            this.radioButtonDesaguesPluviales.Location = new System.Drawing.Point(6, 19);
            this.radioButtonDesaguesPluviales.Name = "radioButtonDesaguesPluviales";
            this.radioButtonDesaguesPluviales.Size = new System.Drawing.Size(118, 17);
            this.radioButtonDesaguesPluviales.TabIndex = 20;
            this.radioButtonDesaguesPluviales.TabStop = true;
            this.radioButtonDesaguesPluviales.Text = "Desagües Pluviales";
            this.radioButtonDesaguesPluviales.UseVisualStyleBackColor = true;
            this.radioButtonDesaguesPluviales.CheckedChanged += new System.EventHandler(this.radioButtonDesaguesPluviales_CheckedChanged);
            // 
            // radioButtonExcavacionCanal
            // 
            this.radioButtonExcavacionCanal.AutoSize = true;
            this.radioButtonExcavacionCanal.Location = new System.Drawing.Point(6, 42);
            this.radioButtonExcavacionCanal.Name = "radioButtonExcavacionCanal";
            this.radioButtonExcavacionCanal.Size = new System.Drawing.Size(126, 17);
            this.radioButtonExcavacionCanal.TabIndex = 21;
            this.radioButtonExcavacionCanal.TabStop = true;
            this.radioButtonExcavacionCanal.Text = "Excavacion de Canal";
            this.radioButtonExcavacionCanal.UseVisualStyleBackColor = true;
            this.radioButtonExcavacionCanal.CheckedChanged += new System.EventHandler(this.radioButtonExcavacionCanal_CheckedChanged);
            // 
            // radioButtonPresas
            // 
            this.radioButtonPresas.AutoSize = true;
            this.radioButtonPresas.Location = new System.Drawing.Point(6, 65);
            this.radioButtonPresas.Name = "radioButtonPresas";
            this.radioButtonPresas.Size = new System.Drawing.Size(57, 17);
            this.radioButtonPresas.TabIndex = 22;
            this.radioButtonPresas.TabStop = true;
            this.radioButtonPresas.Text = "Presas";
            this.radioButtonPresas.UseVisualStyleBackColor = true;
            this.radioButtonPresas.CheckedChanged += new System.EventHandler(this.radioButtonPresas_CheckedChanged);
            // 
            // radioButtonDefensaCostera
            // 
            this.radioButtonDefensaCostera.AutoSize = true;
            this.radioButtonDefensaCostera.Location = new System.Drawing.Point(6, 88);
            this.radioButtonDefensaCostera.Name = "radioButtonDefensaCostera";
            this.radioButtonDefensaCostera.Size = new System.Drawing.Size(98, 17);
            this.radioButtonDefensaCostera.TabIndex = 23;
            this.radioButtonDefensaCostera.TabStop = true;
            this.radioButtonDefensaCostera.Text = "Defensa Costra";
            this.radioButtonDefensaCostera.UseVisualStyleBackColor = true;
            this.radioButtonDefensaCostera.CheckedChanged += new System.EventHandler(this.radioButtonDefensaCostera_CheckedChanged);
            // 
            // radioButtonDefensaPoblacion
            // 
            this.radioButtonDefensaPoblacion.AutoSize = true;
            this.radioButtonDefensaPoblacion.Location = new System.Drawing.Point(6, 111);
            this.radioButtonDefensaPoblacion.Name = "radioButtonDefensaPoblacion";
            this.radioButtonDefensaPoblacion.Size = new System.Drawing.Size(130, 17);
            this.radioButtonDefensaPoblacion.TabIndex = 24;
            this.radioButtonDefensaPoblacion.TabStop = true;
            this.radioButtonDefensaPoblacion.Text = "Defensa de Poblacion";
            this.radioButtonDefensaPoblacion.UseVisualStyleBackColor = true;
            this.radioButtonDefensaPoblacion.CheckedChanged += new System.EventHandler(this.radioButtonDefensaPoblacion_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButtonDesaguesPluviales);
            this.groupBox1.Controls.Add(this.radioButtonDefensaPoblacion);
            this.groupBox1.Controls.Add(this.radioButtonExcavacionCanal);
            this.groupBox1.Controls.Add(this.radioButtonDefensaCostera);
            this.groupBox1.Controls.Add(this.radioButtonPresas);
            this.groupBox1.Location = new System.Drawing.Point(756, 401);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(269, 145);
            this.groupBox1.TabIndex = 25;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Obras";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(26, 466);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(207, 40);
            this.button1.TabIndex = 26;
            this.button1.Text = "Costo financiero anterior";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.Location = new System.Drawing.Point(26, 512);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(207, 40);
            this.button2.TabIndex = 27;
            this.button2.Text = "Nuevo costo financiero";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // maskedViejoCostoFinanciero
            // 
            this.maskedViejoCostoFinanciero.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.maskedViejoCostoFinanciero.Location = new System.Drawing.Point(251, 478);
            this.maskedViejoCostoFinanciero.Name = "maskedViejoCostoFinanciero";
            this.maskedViejoCostoFinanciero.Size = new System.Drawing.Size(100, 22);
            this.maskedViejoCostoFinanciero.TabIndex = 28;
            // 
            // maskedNuevoCostoFinanciero
            // 
            this.maskedNuevoCostoFinanciero.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.maskedNuevoCostoFinanciero.Location = new System.Drawing.Point(251, 522);
            this.maskedNuevoCostoFinanciero.Name = "maskedNuevoCostoFinanciero";
            this.maskedNuevoCostoFinanciero.Size = new System.Drawing.Size(100, 22);
            this.maskedNuevoCostoFinanciero.TabIndex = 29;
            // 
            // buttonTotal
            // 
            this.buttonTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonTotal.Location = new System.Drawing.Point(1038, 499);
            this.buttonTotal.Name = "buttonTotal";
            this.buttonTotal.Size = new System.Drawing.Size(116, 40);
            this.buttonTotal.TabIndex = 30;
            this.buttonTotal.Text = "Total";
            this.buttonTotal.UseVisualStyleBackColor = true;
            this.buttonTotal.Click += new System.EventHandler(this.buttonTotal_Click);
            // 
            // labelCredito
            // 
            this.labelCredito.AutoSize = true;
            this.labelCredito.Font = new System.Drawing.Font("Microsoft YaHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelCredito.Location = new System.Drawing.Point(715, 552);
            this.labelCredito.Name = "labelCredito";
            this.labelCredito.Size = new System.Drawing.Size(439, 19);
            this.labelCredito.TabIndex = 31;
            this.labelCredito.Text = "Creado por: Jonatan Daniel Monti; Email: jonatan.monti@hotmail.com;";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ClientSize = new System.Drawing.Size(1166, 580);
            this.Controls.Add(this.labelCredito);
            this.Controls.Add(this.buttonTotal);
            this.Controls.Add(this.maskedNuevoCostoFinanciero);
            this.Controls.Add(this.maskedViejoCostoFinanciero);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.buttonExportarExcel);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.buttonRutaTXT);
            this.Controls.Add(this.labelRutaTXT);
            this.Controls.Add(this.textBoxRutaTXT);
            this.Controls.Add(this.buttonPruebas);
            this.Controls.Add(this.maskedTextBoxUltimaPagina);
            this.Controls.Add(this.maskedTextBoxPrimeraPagina);
            this.Controls.Add(this.buttonUltimaPagina);
            this.Controls.Add(this.buttonPrimeraPagina);
            this.Controls.Add(this.buttonContinuar);
            this.Controls.Add(this.buttontxt);
            this.Controls.Add(this.buttonRutaPDF);
            this.Controls.Add(this.labelRutaPdf);
            this.Controls.Add(this.textBoxRuta);
            this.Controls.Add(this.labelTitulo);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Form1";
            this.Text = "Calculo en la Variación de Obras Hidráulicas";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelTitulo;
        private System.Windows.Forms.Label labelRutaPdf;
        private System.Windows.Forms.TextBox textBoxRuta;
        private System.Windows.Forms.Button buttonRutaPDF;
        private System.Windows.Forms.Button buttontxt;
        private System.Windows.Forms.Button buttonContinuar;
        private System.Windows.Forms.Button buttonPrimeraPagina;
        private System.Windows.Forms.Button buttonUltimaPagina;
        private System.Windows.Forms.MaskedTextBox maskedTextBoxPrimeraPagina;
        private System.Windows.Forms.MaskedTextBox maskedTextBoxUltimaPagina;
        private System.Windows.Forms.Button buttonPruebas;
        private System.Windows.Forms.TextBox textBoxRutaTXT;
        private System.Windows.Forms.Label labelRutaTXT;
        private System.Windows.Forms.Button buttonRutaTXT;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button buttonExportarExcel;
        private System.Windows.Forms.RadioButton radioButtonDesaguesPluviales;
        private System.Windows.Forms.RadioButton radioButtonExcavacionCanal;
        private System.Windows.Forms.RadioButton radioButtonPresas;
        private System.Windows.Forms.RadioButton radioButtonDefensaCostera;
        private System.Windows.Forms.RadioButton radioButtonDefensaPoblacion;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.MaskedTextBox maskedViejoCostoFinanciero;
        private System.Windows.Forms.MaskedTextBox maskedNuevoCostoFinanciero;
        private System.Windows.Forms.Button buttonTotal;
        private System.Windows.Forms.Label labelCredito;
    }
}

