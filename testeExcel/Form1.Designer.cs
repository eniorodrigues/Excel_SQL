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
            this.SuspendLayout();
            // 
            // buttonAbrir
            // 
            this.buttonAbrir.Location = new System.Drawing.Point(26, 18);
            this.buttonAbrir.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.buttonAbrir.Name = "buttonAbrir";
            this.buttonAbrir.Size = new System.Drawing.Size(210, 88);
            this.buttonAbrir.TabIndex = 9;
            this.buttonAbrir.Text = "Abrir Arquivo";
            this.buttonAbrir.UseVisualStyleBackColor = true;
            this.buttonAbrir.Click += new System.EventHandler(this.buttonAbrir_Click);
            // 
            // lblEndereço
            // 
            this.lblEndereço.AutoSize = true;
            this.lblEndereço.Location = new System.Drawing.Point(25, 264);
            this.lblEndereço.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblEndereço.Name = "lblEndereço";
            this.lblEndereço.Size = new System.Drawing.Size(0, 20);
            this.lblEndereço.TabIndex = 11;
            // 
            // cmbPlanilha
            // 
            this.cmbPlanilha.FormattingEnabled = true;
            this.cmbPlanilha.Location = new System.Drawing.Point(26, 215);
            this.cmbPlanilha.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cmbPlanilha.Name = "cmbPlanilha";
            this.cmbPlanilha.Size = new System.Drawing.Size(442, 28);
            this.cmbPlanilha.TabIndex = 22;
            this.cmbPlanilha.SelectedIndexChanged += new System.EventHandler(this.comboBox2_SelectedIndexChanged);
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
            this.btnCarregar.Location = new System.Drawing.Point(255, 18);
            this.btnCarregar.Name = "btnCarregar";
            this.btnCarregar.Size = new System.Drawing.Size(210, 88);
            this.btnCarregar.TabIndex = 27;
            this.btnCarregar.Text = "Carregar";
            this.btnCarregar.UseVisualStyleBackColor = true;
            this.btnCarregar.Click += new System.EventHandler(this.button1_Click_2);
            // 
            // lblPlanilha
            // 
            this.lblPlanilha.AutoSize = true;
            this.lblPlanilha.Location = new System.Drawing.Point(25, 190);
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
            this.cmbTabela.Location = new System.Drawing.Point(26, 148);
            this.cmbTabela.Name = "cmbTabela";
            this.cmbTabela.Size = new System.Drawing.Size(442, 28);
            this.cmbTabela.TabIndex = 29;
            this.cmbTabela.SelectedIndexChanged += new System.EventHandler(this.cmbTabela_SelectedIndexChanged);
            // 
            // lblTabela
            // 
            this.lblTabela.AutoSize = true;
            this.lblTabela.Location = new System.Drawing.Point(25, 125);
            this.lblTabela.Name = "lblTabela";
            this.lblTabela.Size = new System.Drawing.Size(57, 20);
            this.lblTabela.TabIndex = 30;
            this.lblTabela.Text = "Tabela";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(488, 313);
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
    }
}

