﻿namespace JACA
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
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
            this.btnGerarModelo = new System.Windows.Forms.Button();
            this.lblCarregada = new System.Windows.Forms.Label();
            this.lblTotal = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lblPendencia = new System.Windows.Forms.Label();
            this.lblRepetido = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonAbrir
            // 
            this.buttonAbrir.Enabled = false;
            this.buttonAbrir.Location = new System.Drawing.Point(8, 83);
            this.buttonAbrir.Name = "buttonAbrir";
            this.buttonAbrir.Size = new System.Drawing.Size(369, 43);
            this.buttonAbrir.TabIndex = 3;
            this.buttonAbrir.Text = "Abrir Arquivo Excel";
            this.buttonAbrir.UseVisualStyleBackColor = true;
            this.buttonAbrir.Click += new System.EventHandler(this.buttonAbrir_Click);
            // 
            // lblEndereço
            // 
            this.lblEndereço.AutoSize = true;
            this.lblEndereço.Location = new System.Drawing.Point(11, 335);
            this.lblEndereço.Name = "lblEndereço";
            this.lblEndereço.Size = new System.Drawing.Size(0, 13);
            this.lblEndereço.TabIndex = 11;
            // 
            // cmbPlanilha
            // 
            this.cmbPlanilha.Enabled = false;
            this.cmbPlanilha.FormattingEnabled = true;
            this.cmbPlanilha.Location = new System.Drawing.Point(8, 152);
            this.cmbPlanilha.Name = "cmbPlanilha";
            this.cmbPlanilha.Size = new System.Drawing.Size(371, 21);
            this.cmbPlanilha.TabIndex = 22;
            this.cmbPlanilha.SelectedIndexChanged += new System.EventHandler(this.cmbPlanilha_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(428, 250);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 13);
            this.label1.TabIndex = 23;
            // 
            // btnCarregar
            // 
            this.btnCarregar.Enabled = false;
            this.btnCarregar.Location = new System.Drawing.Point(8, 236);
            this.btnCarregar.Margin = new System.Windows.Forms.Padding(2);
            this.btnCarregar.Name = "btnCarregar";
            this.btnCarregar.Size = new System.Drawing.Size(371, 36);
            this.btnCarregar.TabIndex = 27;
            this.btnCarregar.Text = "Carregar";
            this.btnCarregar.UseVisualStyleBackColor = true;
            this.btnCarregar.Click += new System.EventHandler(this.btnCarregar_Click);
            // 
            // lblPlanilha
            // 
            this.lblPlanilha.AutoSize = true;
            this.lblPlanilha.Location = new System.Drawing.Point(5, 136);
            this.lblPlanilha.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblPlanilha.Name = "lblPlanilha";
            this.lblPlanilha.Size = new System.Drawing.Size(73, 13);
            this.lblPlanilha.TabIndex = 28;
            this.lblPlanilha.Text = "Planilha Excel";
            // 
            // cmbTabela
            // 
            this.cmbTabela.Enabled = false;
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
            "D_Vendas_Itens",
            "D_PIC"});
            this.cmbTabela.Location = new System.Drawing.Point(8, 214);
            this.cmbTabela.Margin = new System.Windows.Forms.Padding(2);
            this.cmbTabela.Name = "cmbTabela";
            this.cmbTabela.Size = new System.Drawing.Size(371, 21);
            this.cmbTabela.TabIndex = 29;
            this.cmbTabela.SelectedIndexChanged += new System.EventHandler(this.cmbTabela_SelectedIndexChanged);
            // 
            // lblTabela
            // 
            this.lblTabela.AutoSize = true;
            this.lblTabela.Location = new System.Drawing.Point(2, 199);
            this.lblTabela.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblTabela.Name = "lblTabela";
            this.lblTabela.Size = new System.Drawing.Size(64, 13);
            this.lblTabela.TabIndex = 30;
            this.lblTabela.Text = "Tabela SQL";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(10, 38);
            this.label7.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(81, 13);
            this.label7.TabIndex = 35;
            this.label7.Text = "Base de dados:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(43, 16);
            this.label6.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(49, 13);
            this.label6.TabIndex = 34;
            this.label6.Text = "Servidor:";
            // 
            // comboBoxBase
            // 
            this.comboBoxBase.Enabled = false;
            this.comboBoxBase.FormattingEnabled = true;
            this.comboBoxBase.Location = new System.Drawing.Point(93, 33);
            this.comboBoxBase.Margin = new System.Windows.Forms.Padding(2);
            this.comboBoxBase.Name = "comboBoxBase";
            this.comboBoxBase.Size = new System.Drawing.Size(285, 21);
            this.comboBoxBase.Sorted = true;
            this.comboBoxBase.TabIndex = 2;
            this.comboBoxBase.SelectedIndexChanged += new System.EventHandler(this.comboBoxBase_SelectedIndexChanged);
            // 
            // comboBoxServidor
            // 
            this.comboBoxServidor.FormattingEnabled = true;
            this.comboBoxServidor.Location = new System.Drawing.Point(93, 11);
            this.comboBoxServidor.Margin = new System.Windows.Forms.Padding(2);
            this.comboBoxServidor.Name = "comboBoxServidor";
            this.comboBoxServidor.Size = new System.Drawing.Size(285, 21);
            this.comboBoxServidor.TabIndex = 1;
            this.comboBoxServidor.SelectedIndexChanged += new System.EventHandler(this.comboBoxServidor_SelectedIndexChanged);
            this.comboBoxServidor.Enter += new System.EventHandler(this.comboBoxServidor_Enter);
            // 
            // btnGerarModelo
            // 
            this.btnGerarModelo.Location = new System.Drawing.Point(8, 55);
            this.btnGerarModelo.Margin = new System.Windows.Forms.Padding(2);
            this.btnGerarModelo.Name = "btnGerarModelo";
            this.btnGerarModelo.Size = new System.Drawing.Size(125, 23);
            this.btnGerarModelo.TabIndex = 36;
            this.btnGerarModelo.Text = "Gerar Modelo de Excel ";
            this.btnGerarModelo.UseVisualStyleBackColor = true;
            this.btnGerarModelo.Click += new System.EventHandler(this.btnGerarModelo_Click);
            // 
            // lblCarregada
            // 
            this.lblCarregada.AutoSize = true;
            this.lblCarregada.Location = new System.Drawing.Point(180, 292);
            this.lblCarregada.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblCarregada.Name = "lblCarregada";
            this.lblCarregada.Size = new System.Drawing.Size(13, 13);
            this.lblCarregada.TabIndex = 37;
            this.lblCarregada.Text = "0";
            // 
            // lblTotal
            // 
            this.lblTotal.AutoSize = true;
            this.lblTotal.Location = new System.Drawing.Point(316, 136);
            this.lblTotal.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblTotal.Name = "lblTotal";
            this.lblTotal.Size = new System.Drawing.Size(13, 13);
            this.lblTotal.TabIndex = 37;
            this.lblTotal.Text = "0";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(230, 136);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(83, 13);
            this.label4.TabIndex = 38;
            this.label4.Text = "Total de Linhas:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(5, 292);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(98, 13);
            this.label5.TabIndex = 38;
            this.label5.Text = "Linhas Carregadas:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(5, 305);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(169, 13);
            this.label2.TabIndex = 38;
            this.label2.Text = "Pendência de Layout e Conteúdo:";
            // 
            // lblPendencia
            // 
            this.lblPendencia.AutoSize = true;
            this.lblPendencia.Location = new System.Drawing.Point(180, 305);
            this.lblPendencia.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblPendencia.Name = "lblPendencia";
            this.lblPendencia.Size = new System.Drawing.Size(13, 13);
            this.lblPendencia.TabIndex = 37;
            this.lblPendencia.Text = "0";
            // 
            // lblRepetido
            // 
            this.lblRepetido.AutoSize = true;
            this.lblRepetido.Location = new System.Drawing.Point(180, 318);
            this.lblRepetido.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblRepetido.Name = "lblRepetido";
            this.lblRepetido.Size = new System.Drawing.Size(13, 13);
            this.lblRepetido.TabIndex = 37;
            this.lblRepetido.Text = "0";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(5, 318);
            this.label8.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(58, 13);
            this.label8.TabIndex = 38;
            this.label8.Text = "Repetidos:";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(390, 359);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.lblRepetido);
            this.Controls.Add(this.lblTotal);
            this.Controls.Add(this.lblPendencia);
            this.Controls.Add(this.lblCarregada);
            this.Controls.Add(this.btnGerarModelo);
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
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.ShowIcon = false;
            this.Text = "Tool Versão de Teste 1.6";
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
        private System.Windows.Forms.Button btnGerarModelo;
        private System.Windows.Forms.Label lblCarregada;
        private System.Windows.Forms.Label lblTotal;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblPendencia;
        private System.Windows.Forms.Label lblRepetido;
        private System.Windows.Forms.Label label8;
    }
}

