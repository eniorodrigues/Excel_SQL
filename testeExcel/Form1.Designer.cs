namespace testeExcel
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
            this.buttonAbrir = new System.Windows.Forms.Button();
            this.lblEndereço = new System.Windows.Forms.Label();
            this.cmbPlanilha = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnCarregar = new System.Windows.Forms.Button();
            this.lblPlanilha = new System.Windows.Forms.Label();
            this.cmbTabela = new System.Windows.Forms.ComboBox();
            this.lblTabela = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.comboBoxBase = new System.Windows.Forms.ComboBox();
            this.comboBoxServidor = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonAbrir
            // 
            this.buttonAbrir.Location = new System.Drawing.Point(10, 92);
            this.buttonAbrir.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.buttonAbrir.Name = "buttonAbrir";
            this.buttonAbrir.Size = new System.Drawing.Size(441, 66);
            this.buttonAbrir.TabIndex = 3;
            this.buttonAbrir.Text = "Abrir Arquivo Excel";
            this.buttonAbrir.UseVisualStyleBackColor = true;
            this.buttonAbrir.Click += new System.EventHandler(this.buttonAbrir_Click);
            // 
            // lblEndereço
            // 
            this.lblEndereço.AutoSize = true;
            this.lblEndereço.Location = new System.Drawing.Point(5, 253);
            this.lblEndereço.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblEndereço.Name = "lblEndereço";
            this.lblEndereço.Size = new System.Drawing.Size(0, 20);
            this.lblEndereço.TabIndex = 11;
            // 
            // cmbPlanilha
            // 
            this.cmbPlanilha.FormattingEnabled = true;
            this.cmbPlanilha.Location = new System.Drawing.Point(10, 204);
            this.cmbPlanilha.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cmbPlanilha.Name = "cmbPlanilha";
            this.cmbPlanilha.Size = new System.Drawing.Size(442, 28);
            this.cmbPlanilha.TabIndex = 22;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(642, 384);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 20);
            this.label1.TabIndex = 23;
            // 
            // btnCarregar
            // 
            this.btnCarregar.Location = new System.Drawing.Point(163, 338);
            this.btnCarregar.Name = "btnCarregar";
            this.btnCarregar.Size = new System.Drawing.Size(288, 55);
            this.btnCarregar.TabIndex = 27;
            this.btnCarregar.Text = "Carregar";
            this.btnCarregar.UseVisualStyleBackColor = true;
            this.btnCarregar.Click += new System.EventHandler(this.button1_Click_2);
            // 
            // lblPlanilha
            // 
            this.lblPlanilha.AutoSize = true;
            this.lblPlanilha.Location = new System.Drawing.Point(9, 178);
            this.lblPlanilha.Name = "lblPlanilha";
            this.lblPlanilha.Size = new System.Drawing.Size(106, 20);
            this.lblPlanilha.TabIndex = 28;
            this.lblPlanilha.Text = "Planilha Excel";
            // 
            // cmbTabela
            // 
            this.cmbTabela.FormattingEnabled = true;
            this.cmbTabela.Items.AddRange(new object[] {
            "D_Clientes",
            "D_Compras",
            "D_Custo_Medio",
            "D_Fornecedores",
            "D_Insumo_Produto",
            "D_Inventario_Carga",
            "D_Produtos",
            "D_Relacao_Carga",
            "D_Vendas_Itens"});
            this.cmbTabela.Location = new System.Drawing.Point(7, 296);
            this.cmbTabela.Name = "cmbTabela";
            this.cmbTabela.Size = new System.Drawing.Size(442, 28);
            this.cmbTabela.TabIndex = 29;
            // 
            // lblTabela
            // 
            this.lblTabela.AutoSize = true;
            this.lblTabela.Location = new System.Drawing.Point(6, 273);
            this.lblTabela.Name = "lblTabela";
            this.lblTabela.Size = new System.Drawing.Size(93, 20);
            this.lblTabela.TabIndex = 30;
            this.lblTabela.Text = "Tabela SQL";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(17, 54);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(120, 20);
            this.label7.TabIndex = 35;
            this.label7.Text = "Base de dados:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(66, 20);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(71, 20);
            this.label6.TabIndex = 34;
            this.label6.Text = "Servidor:";
            // 
            // comboBoxBase
            // 
            this.comboBoxBase.FormattingEnabled = true;
            this.comboBoxBase.Location = new System.Drawing.Point(141, 46);
            this.comboBoxBase.Name = "comboBoxBase";
            this.comboBoxBase.Size = new System.Drawing.Size(311, 28);
            this.comboBoxBase.TabIndex = 2;
            this.comboBoxBase.SelectedIndexChanged += new System.EventHandler(this.comboBoxBase_SelectedIndexChanged);
            // 
            // comboBoxServidor
            // 
            this.comboBoxServidor.FormattingEnabled = true;
            this.comboBoxServidor.Location = new System.Drawing.Point(141, 12);
            this.comboBoxServidor.Name = "comboBoxServidor";
            this.comboBoxServidor.Size = new System.Drawing.Size(311, 28);
            this.comboBoxServidor.TabIndex = 1;
            this.comboBoxServidor.SelectedIndexChanged += new System.EventHandler(this.comboBoxServidor_SelectedIndexChanged);
            this.comboBoxServidor.Enter += new System.EventHandler(this.comboBoxServidor_Enter);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(10, 338);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(147, 55);
            this.button1.TabIndex = 36;
            this.button1.Text = "Gerar Excel Modelo";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_5);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(468, 405);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.comboBoxBase);
            this.Controls.Add(this.comboBoxServidor);
            this.Controls.Add(this.lblTabela);
            this.Controls.Add(this.cmbTabela);
            this.Controls.Add(this.lblPlanilha);
            this.Controls.Add(this.btnCarregar);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbPlanilha);
            this.Controls.Add(this.lblEndereço);
            this.Controls.Add(this.buttonAbrir);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "Form1";
            this.ShowIcon = false;
            this.Text = "Tool";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button buttonAbrir;
        private System.Windows.Forms.Label lblEndereço;
        private System.Windows.Forms.ComboBox cmbPlanilha;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnCarregar;
        private System.Windows.Forms.Label lblPlanilha;
        private System.Windows.Forms.ComboBox cmbTabela;
        private System.Windows.Forms.Label lblTabela;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox comboBoxBase;
        private System.Windows.Forms.ComboBox comboBoxServidor;
        private System.Windows.Forms.Button button1;
    }
}

