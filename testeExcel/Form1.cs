using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
//using OfficeOpenXml;
using System.Linq;
using System.Text;
//using ExcelIt = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace testeExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        public static string path;
        public static string excelConnectionString;
        public string[] files;
        public string conexao;
        public string baseDeDados;
        public string tabela;
        public string caminho;
        public string directoryPath;
        private static Excel.Application MyApp = null;
        public List<string> filesAdionado = new List<string>();
        public List<string> colunas = new List<string>();
        public List<string> colunasCreate = new List<string>();
        public string tipoArquivo;
        Stream myStream = null;
        StringBuilder camposDataGrid = new StringBuilder();
        public List<String> itemsDataGrid = new List<String>();
        public SqlConnection conn = null;
        public SqlConnection conn1 = null;
        public bool checado;
        String databaseName;

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void buttonAbrir_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "C:\\";
            openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.Multiselect = true;
            cmbPlanilha.Text = "";
            lblTotal.Text = "0";
            lblCarregada.Text = "0";
            lblPendencia.Text = "0";
            lblRepetido.Text = "0";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        using (myStream)
                        {
                            caminho = openFileDialog1.FileName;
                            directoryPath = Path.GetDirectoryName(openFileDialog1.FileName);
                            files = (openFileDialog1.SafeFileNames);
                            carregaLinhas();
                        }
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        public void carregaLinhas()
        {
            cmbPlanilha.Items.Clear();
            lblEndereço.Text = caminho;
            MyApp = new Excel.Application();
            excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + caminho + ";Extended Properties=Excel 12.0;";
            MyApp.Workbooks.Add("");
            MyApp.Workbooks.Add(caminho);

            for (int i = 1; i <= MyApp.Workbooks[2].Worksheets.Count; i++)
            {
                cmbPlanilha.Items.Add(MyApp.Workbooks[2].Worksheets[i].Name);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        public void Vendas()
        {
            int linhaRegistro = 1;
            int registroInconsistente = 0;
            int registroConsistente = 0;
            bool penLayout = false;
            string filePath = caminho;

            try
            {
                FileInfo existingFile = new FileInfo(filePath);
                ExcelPackage package = new ExcelPackage(existingFile);
                ExcelWorksheet workSheet = package.Workbook.Worksheets[cmbPlanilha.SelectedIndex + 1];
                StringBuilder conteudo = new StringBuilder();
                SqlCommand cmd = conn.CreateCommand();

                ///// temporario
                //conn = new SqlConnection("Data Source=BRCAENRODRIGUES\\SQLEXPRESS01; Integrated Security=True; Initial Catalog=LAMPADA");
                //filePath = @"C:\Base\Vendas_Doosan_Jan_Jun_2019.xlsx";
                //FileInfo existingFile = new FileInfo(filePath);
                //ExcelPackage package = new ExcelPackage(existingFile);
                //ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
                //StringBuilder conteudo = new StringBuilder();
                //SqlCommand cmd = conn.CreateCommand();

                lblTotal.Text = workSheet.Dimension.End.Row.ToString();
                lblTotal.Refresh();
                //MessageBox.Show(workSheet.Dimension.End.Row.ToString());

                for (int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
                {
                    lblCarregada.Text = i.ToString();
                    lblCarregada.Refresh();

                    penLayout = false;
                    for (int j = workSheet.Dimension.Start.Column; j <= 25; j++)
                    {
                        if (j == 25)
                        {
                            if (workSheet.Cells[i, j].Value == null)
                            {
                                conteudo.Append(" '', '" + linhaRegistro + "', ");
                                conteudo.Append(" " + pegarID("D_Vendas_Itens") + " ");
                            }
                            else
                            {
                                conteudo.Append(" '" + workSheet.Cells[i, j].Value.ToString() + "' , '" + linhaRegistro + "', ");
                                conteudo.Append(" " + pegarID("D_Vendas_Itens") + " ");
                            }
                        }
                        else
                        {
                            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US")
                            {
                                DateTimeFormat = { YearMonthPattern = "yyyy-mm-dd" }
                            };
                            if (j == 1)
                            {
                                conteudo.Clear();
                                conteudo.Append(" INSERT INTO[dbo].[D_Vendas_Itens] " +
                                                  "([Vnd_Cli_ID] " +
                                                  ",[Vnd_NF_ID]" +
                                                  ",[Vnd_NF_Serie]" +
                                                  ",[Vnd_Cod_Divisao]" +
                                                  ",[Vnd_CFOP]" +
                                                  ",[Vnd_Dt_Emissao]" +
                                                  ",[Vnd_DT_Vencimento]" +
                                                  ",[Vnd_Dias]" +
                                                  ",[Vnd_Item]" +
                                                  ",[Vnd_Pro_id]" +
                                                  ",[Vnd_Qtde]" +
                                                  ",[Vnd_Vl_Nota]" +
                                                  ",[Vnd_Desconto]" +
                                                  ",[Vnd_ICMS]" +
                                                  ",[Vnd_PIS]" +
                                                  ",[Vnd_COFINS]" +
                                                  ",[Vnd_ISS]" +
                                                  ",[Vnd_Comissao]" +
                                                  ",[Vnd_Frete]" +
                                                  ",[Vnd_Seguro]" +
                                                  ",[Vnd_Dt_Embarque]" +
                                                  ",[Vnd_Cod_Moeda]" +
                                                  ",[Vnd_Vl_Moeda]" +
                                                  ",[Vnd_Custo] " +
                                                  ",[Vnd_RE] " +
                                                  ",[Vnd_CNPJ] " +
                                                  ",[Lin_Origem_ID] " +
                                                  ",[Arq_Origem_ID]) " +
                                                  " VALUES ( ");
                                if (workSheet.Cells[i, j].Value == null)
                                {
                                    VendasPenLayout(i);
                                    j = workSheet.Dimension.End.Column;
                                    penLayout = true;
                                }
                                else
                                {
                                    conteudo.Append("'" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "', ");
                                }
                            }
                            else if ((j == 2 || j == 5 || j == 6 || j == 10 || j == 11 || j == 12) && (workSheet.Cells[i, j].Value == null))
                            {
                                VendasPenLayout(i);
                                penLayout = true;
                            }
                            else if (j == 9 && (workSheet.Cells[i, j].Value == null || workSheet.Cells[i, j].Value.ToString() == ""))
                            {
                                conteudo.Append(" '', ");
                            }
                            else if ((j == 6 || j == 7 || j == 21) && workSheet.Cells[i, j].Value != null && workSheet.Cells[i, j].Value.GetType().Name.ToString() == "DateTime")
                            {
                                conteudo.Append("'" + workSheet.Cells[i, j].Value.ToString() + "', ");
                            }
                            else if ((j == 6 || j == 7 || j == 21) && workSheet.Cells[i, j].Value != null && workSheet.Cells[i, j].Value.GetType().Name.ToString() == "Double")
                            {
                                DateTime dt = DateTime.FromOADate(Convert.ToInt64(workSheet.Cells[i, j].Value));
                                conteudo.Append("'" + dt + "', ");
                            }
                            else if ((j == 6 || j == 7 || j == 21) && workSheet.Cells[i, j].Value != null)
                            {
                                DateTime dt = DateTime.ParseExact(workSheet.Cells[i, j].Value.ToString(), "dd/MM/yyyy", null);
                                conteudo.Append("'" + dt + "', ");
                            }
                            else if ((j == 6 || j == 7 || j == 21) && workSheet.Cells[i, j].Value != null)
                            {
                                DateTime dt = DateTime.ParseExact(workSheet.Cells[i, j].Value.ToString(), "dd/MM/yyyy", null);
                                conteudo.Append("'" + dt + "', ");
                            }
                            else if ((j == 6 || j == 7 || j == 8 || j == 21 || j == 22) && workSheet.Cells[i, j].Value == null)
                            {
                                conteudo.Append(" '', ");
                            }
                            ////caso número seja nulo colocar zero
                            else if ((j == 9 || j == 11 || j == 12 || j == 13 || j == 14 || j == 15 || j == 16 || j == 17 || j == 18 || j == 19 || j == 20 || j == 23 || j == 24) && (workSheet.Cells[i, j].Value == null))
                            {
                                conteudo.Append(" " + 0 + ", ");
                            }
                            //caso o que não é numero esteja em branco colocar texto branco
                            else if ((j == 9 || j == 11 || j == 12 || j == 13 || j == 14 || j == 15 || j == 16 || j == 17 || j == 18 || j == 19 || j == 20 || j == 23 || j == 24) && (workSheet.Cells[i, j].Value.ToString().Equals("")))
                            {
                                conteudo.Append(" " + 0 + ", ");
                            }
                            //caso número tirar  aspas simples
                            else if ((j == 11 || j == 12 || j == 13 || j == 14 || j == 15 || j == 16 || j == 17 || j == 18 || j == 19 || j == 20 || j == 23 || j == 24) && (workSheet.Cells[i, j].Value != null))
                            {
                                conteudo.Append(" " + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + ", ");
                            }
                            else if ((j == 4 || j == 3) && (workSheet.Cells[i, j].Value == null))
                            {
                                conteudo.Append(" '', ");
                            }
                            else if (workSheet.Cells[i, j].Value != null || workSheet.Cells[i, j].Value.ToString() != "")
                            {
                                if ((workSheet.Cells[i, j].Value.GetType().Name.ToString()) == "DateTime")
                                {
                                    workSheet.Cells[i, j].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.YearMonthPattern;
                                    conteudo.Append("'" + workSheet.Cells[i, j].Value.ToString() + "', ");
                                }
                                else
                                {
                                    conteudo.Append("'" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "', ");
                                }
                            }
                            else if (workSheet.Cells[i, j].Value == null)
                            {
                                conteudo.Append(" '', ");
                            }
                        }
                    }
                    if (i == workSheet.Dimension.End.Row)
                    {
                        conteudo.Append(" ) ");
                    }
                    else
                    {
                        conteudo.Append(")");
                        conteudo.Append(Environment.NewLine);
                    }
                    if (penLayout == true)
                    {
                        registroInconsistente++;
                    }
                    else
                    {
                        registroConsistente++;
                    }

                    if (penLayout == false)
                    {
                        cmd = conn.CreateCommand();
                        // Clipboard.SetText(conteudo.ToString());

                        if(conn.State.ToString() == "Closed")
                        {
                        conn.Open();
                        }
                         
                        cmd.CommandText = conteudo.ToString();
                        SqlTransaction trE = null;
                        trE = conn.BeginTransaction();
                        cmd.Transaction = trE;
                        cmd.ExecuteNonQuery();
                        trE.Commit();
                        conteudo.Clear();
                    }
                }
                package.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                SqlCommand cmdArquivoCarregado = conn.CreateCommand();
                cmdArquivoCarregado.CommandText =
                    " declare @tabela varchar(max) = 'D_Vendas_Itens';" +
                    " if (select count(arq_id) from S_ArquivoCarregado where Arq_Tabela = @tabela) = 0" +
                    " insert into S_ArquivoCarregado" +
                    " (Arq_ID, Arq_Nome, Arq_Tabela, Arq_Mensagem, Arq_DataCarga, Arq_Quantidade, Arq_Login)" +
                    " values(1, '" + caminho + "', @tabela, 'Carga efetuada com sucesso.'," +
                    " GETDATE(), " + registroConsistente.ToString() + ", REPLACE(SUSER_NAME(), 'ATRAME\\',''))" +
                    " else" +
                    " insert into S_ArquivoCarregado" +
                    " (Arq_ID, Arq_Nome, Arq_Tabela, Arq_Mensagem, Arq_DataCarga, Arq_Quantidade, Arq_Login)" +
                    " values('" + +pegarID("D_Vendas_Itens") + "', '" + caminho + "', @tabela, 'Carga efetuada com sucesso.'," +
                    " GETDATE(), " + registroConsistente.ToString() + ", REPLACE(SUSER_NAME(), 'ATRAME\\',''))";
                conn.Open();

                SqlTransaction trA = null;
                trA = conn.BeginTransaction();
                cmdArquivoCarregado.Transaction = trA;
                cmdArquivoCarregado.ExecuteNonQuery();
                trA.Commit();
                conn.Close();

                MessageBox.Show(new Form { TopMost = true }, "Carregamento de " + registroConsistente.ToString() + " registros de vendas");
            }
        }

        public void Compras()
        {
            string filePath = caminho;
            int linha = 0;
            int linhaRegistro = 1;
            int registroInconsistente = 0;
            int registroConsistente = 0;
            int registroRepetido = 0;
            bool penLayout = false;

            try
            {
                linha = 1;
                FileInfo existingFile = new FileInfo(filePath);
                ExcelPackage package = new ExcelPackage(existingFile);
                ExcelWorksheet workSheet = package.Workbook.Worksheets[cmbPlanilha.SelectedIndex + 1];
                StringBuilder conteudo = new StringBuilder();
                SqlCommand cmd = conn.CreateCommand();

                lblTotal.Text = workSheet.Dimension.End.Row.ToString();
                lblTotal.Refresh();

                ///// temporario
                //conn = new SqlConnection("Data Source=BRCAENRODRIGUES\\SQLEXPRESS01; Integrated Security=True; Initial Catalog=LAMPADA");
                //filePath = @"C:\Base\Compras_Doosan_Jan_Jun_2019.xlsx";
                //FileInfo existingFile = new FileInfo(filePath);
                //ExcelPackage package = new ExcelPackage(existingFile);
                //ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
                //StringBuilder conteudo = new StringBuilder();
                //SqlCommand cmd = conn.CreateCommand();

                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US")
                {
                    DateTimeFormat = { YearMonthPattern = "yyyy-mm-dd" }
                };

                for (int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
                {
                    lblCarregada.Text = registroConsistente.ToString();
                    lblCarregada.Refresh();

                    lblPendencia.Text = registroInconsistente.ToString();
                    lblPendencia.Refresh();

                    penLayout = false;
                    for (int j = workSheet.Dimension.Start.Column; j <= 30; j++)
                    {
                        if (j == workSheet.Dimension.End.Column)
                        {
                            conteudo.Append(workSheet.Cells[i, j].Value == null || workSheet.Cells[i, j].Value.ToString() == "" ? " '', '" + linhaRegistro + "', " : " '" + workSheet.Cells[i, j].Value.ToString() + "' , '" + linhaRegistro + "', ");
                            conteudo.Append(" " + pegarID("D_Compras") + "  ");
                        }
                        else
                        {
                            if (j == 1)
                            {
                                conteudo.Clear();
                                conteudo.Append(" INSERT INTO[dbo].[D_Compras] " +
                                                "([Cmp_Pro_ID] " +
                                                ",[Cmp_Cod_Divisao]" +
                                                ",[Cmp_For_ID]" +
                                                ",[Cmp_Lanc_Cont]" +
                                                ",[Cmp_Fat_Coml]" +
                                                ",[Cmp_BL_DT]" +
                                                ",[Cmp_DI_ID]" +
                                                ",[Cmp_DI_DT_Emissao]" +
                                                ",[Cmp_NF_Entrada]" +
                                                ",[Cmp_NF_Serie]" +
                                                ",[Cmp_NF_DT]" +
                                                ",[Cmp_CFOP]" +
                                                ",[Cmp_DI_DT_Vencimento]" +
                                                ",[Cmp_DI_Dias]" +
                                                ",[Cmp_Qtde]" +
                                                ",[Cmp_Valor_Fob]" +
                                                ",[Cmp_Cod_Moeda]" +
                                                ",[Cmp_Vl_Frete_Moeda]" +
                                                ",[Cmp_VL_Seguro_Moeda]" +
                                                ",[Cmp_Cod_Moeda_Frete]" +
                                                ",[Cmp_Cod_Moeda_Seguro]" +
                                                ",[Cmp_Imposto_Import]" +
                                                ",[Cmp_ICMS]" +
                                                ",[Cmp_PIS]" +
                                                ",[Cmp_COFINS]" +
                                                ",[Cmp_Und_Id]" +
                                                ",[Cmp_CNPJ]" +
                                                ",[Cmp_Incoterm]" +
                                                ",[Cmp_For_id_Seguro]" +
                                                ",[Cmp_For_id_Frete]" +
                                                ",[Lin_Origem_ID]" +
                                                ",[Arq_Origem_ID])" +
                                                " VALUES ( ");
                                if (workSheet.Cells[i, j].Value == null)
                                {
                                    ComprasPenLayout(i);
                                    j = workSheet.Dimension.End.Column;
                                    penLayout = true;
                                }
                                else
                                {
                                    conteudo.Append("'" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "', ");
                                }
                            }
                            else if ((j == 15 || j == 16) && (workSheet.Cells[i, j].Value == null))
                            {
                                if (workSheet.Cells[i, j].Value == null)
                                {
                                    conteudo.Replace("[D_Compras]", "[D_Compras_Inconsistencias]");
                                    conteudo.Append(" '', ");
                                }
                                else if (workSheet.Cells[i, j].Value.ToString() == "")
                                {
                                    conteudo.Replace("[D_Compras]", "[D_Compras_Inconsistencias]");
                                    conteudo.Append(" '', ");
                                }
                            }
                            else if ((j == 6 || j == 8 || j == 11 || j == 13) && workSheet.Cells[i, j].Value != null && workSheet.Cells[i, j].Value.GetType().Name.ToString() == "DateTime")
                            {
                                conteudo.Append(" '" + workSheet.Cells[i, j].Value.ToString().Replace('.', '/') + "', ");
                            }
                            else if ((j == 6 || j == 8 || j == 11 || j == 13) && workSheet.Cells[i, j].Value != null && workSheet.Cells[i, j].Value.GetType().Name.ToString() == "Double")
                            {
                                DateTime dt = DateTime.FromOADate(Convert.ToInt64(workSheet.Cells[i, j].Value.ToString().Replace('.', '/')));
                                conteudo.Append(" '" + dt + "', ");
                            }
                            else if ((j == 6 || j == 8 || j == 11 || j == 13) && workSheet.Cells[i, j].Value != null && workSheet.Cells[i, j].Value.GetType().Name.ToString() == "String")
                            {
                                DateTime dt = DateTime.ParseExact(workSheet.Cells[i, j].Value.ToString().Replace('.', '/'), "dd/MM/yyyy", null);
                                conteudo.Append("'" + dt + "', ");

                                //DateTime dt = DateTime.FromOADate(Convert.ToInt64(workSheet.Cells[i, j].Value.ToString().Replace('.', '/')));
                                //conteudo.Append(" '" + dt + "', ");
                            }
                            else if ((j == 6 || j == 8 || j == 11 || j == 13) && workSheet.Cells[i, j].Value != null)
                            {
                                conteudo.Append(" '" + workSheet.Cells[i, j].Value.ToString().Replace('.', '/') + "', ");
                            }
                            //caso o que não é numero esteja em nulo colocar texto branco
                            else if ((j == 23 || j == 24 || j == 25 || j == 19 || j == 18 || j == 15 || j == 16 || j == 22) && (workSheet.Cells[i, j].Value == null))
                            {
                                conteudo.Append(" " + 0 + ", ");
                            }
                            //caso o que não é numero esteja em nulo colocar texto branco
                            else if ((j == 1 || j == 3 || j == 9 || j == 12) && (workSheet.Cells[i, j].Value == null))
                            {
                                ComprasPenLayout(i);
                                penLayout = true;
                            }
                            //caso número seja branco colocar zero
                            else if ((j == 23 || j == 24 || j == 25 || j == 19 || j == 18 || j == 15 || j == 16 || j == 22) && (workSheet.Cells[linha, j].Value.ToString() == ""))
                            {
                                conteudo.Append(" " + 0 + ", ");
                            }
                            //caso número tirar  aspas simples
                            else if ((j == 23 || j == 24 || j == 25 || j == 19 || j == 18 || j == 15 || j == 16 || j == 22) && (workSheet.Cells[i, j].Value != null))
                            {
                                conteudo.Append(" " + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + ", ");
                            }
                            else
                            {
                                if ((workSheet.Cells[i, j].Value == null ? " NULL " : workSheet.Cells[i, j].Value.GetType().Name.ToString()) == "DateTime")
                                {
                                    conteudo.Append(" '" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "', ");
                                }
                                else
                                {
                                    conteudo.Append(workSheet.Cells[i, j].Value == null ? " NULL, " : "'" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "', ");
                                }
                            }
                        }
                    }
                    if (i == workSheet.Dimension.End.Row)
                    {
                        conteudo.Append(" ) ");
                    }
                    else
                    {
                        conteudo.Append(")");
                        conteudo.Append(Environment.NewLine);
                    }
                    if (penLayout == true)
                    {
                        registroInconsistente++;
                    }
                    else
                    {
                        registroConsistente++;
                    }
                    linhaRegistro++;
                    if (penLayout == false)
                    {
                        cmd = conn.CreateCommand();
                          Clipboard.SetText(conteudo.ToString());
                        conn.Open();
                        cmd.CommandText = conteudo.ToString();
                        SqlTransaction trE = null;
                        trE = conn.BeginTransaction();
                        cmd.Transaction = trE;
                        cmd.ExecuteNonQuery();
                        trE.Commit();
                        conteudo.Clear();
                    }
                }
                package.Dispose();

                MessageBox.Show(new Form { TopMost = true }, "Carregamento de " + registroConsistente.ToString() + " registros de compras concluído!");

                SqlCommand cmdArquivoCarregado = conn.CreateCommand();
                cmdArquivoCarregado.CommandText =
                " declare @tabela varchar(max) = 'D_Compras';" +
                " if (select count(arq_id) from S_ArquivoCarregado where Arq_Tabela = @tabela) = 0" +
                " insert into S_ArquivoCarregado" +
                " (Arq_ID, Arq_Nome, Arq_Tabela, Arq_Mensagem, Arq_DataCarga, Arq_Quantidade, Arq_Login)" +
                " values(1, '" + caminho + "', @tabela, 'Carga efetuada com sucesso.'," +
                " GETDATE(), " + registroConsistente.ToString() + ", REPLACE(SUSER_NAME(), 'ATRAME\\',''))" +
                " else" +
                " insert into S_ArquivoCarregado" +
                " (Arq_ID, Arq_Nome, Arq_Tabela, Arq_Mensagem, Arq_DataCarga, Arq_Quantidade, Arq_Login)" +
                " values(" + pegarID("D_Compras") + ", '" + caminho + "', @tabela, 'Carga efetuada com sucesso.'," +
                " GETDATE(), " + registroConsistente.ToString() + ", REPLACE(SUSER_NAME(), 'ATRAME\\',''))";

                conn.Open();
                SqlTransaction trA = null;
                trA = conn.BeginTransaction();
                cmdArquivoCarregado.Transaction = trA;
                cmdArquivoCarregado.ExecuteNonQuery();
                trA.Commit();
                conn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        public void ComprasPenLayout(int linha)
        {
            string filePath = caminho;
            try
            {
                /////temporario
                /////filePath = @"C:\Base\Compras_Doosan_Jan_Jun_2019.xlsx";

                FileInfo existingFile = new FileInfo(filePath);
                ExcelPackage package = new ExcelPackage(existingFile);
                ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
                StringBuilder conteudo = new StringBuilder();

                for (int j = workSheet.Dimension.Start.Column; j <= workSheet.Dimension.End.Column; j++)
                {
                    if (j == 1)
                    {
                        conteudo.Append(" INSERT INTO[dbo].[A_PendenciaLayout] " +
                                            " ([pen_Linha], " +
                                            " [pen_Tabela]," +
                                            " [pen_Campo]," +
                                            " [pen_Posi]," +
                                            " [pen_Tam]," +
                                            " [pen_Erro]," +
                                            " [pen_Registro]," +
                                            " [pen_Arq_Origem])" +
                                            " VALUES ( ");
                        if (workSheet.Cells[linha, j].Value == null)
                        {
                            int registro = linha - 1;
                            conteudo.Append(" " + registro + ", 'D_Compras', 'Cmp_Pro_id', 0, 0, 'Campo [Código do Produto] é obrigatório', '");
                        }
                        else
                        {
                            int registro = linha - 1;
                            conteudo.Append(" " + registro + ", 'D_Compras', 'Cmp_Pro_id', 0, 0, 'Campo [Código do Produto] é obrigatório', '" + workSheet.Cells[linha, j].Value.ToString() + " ");
                        }
                    }
                    else if (j == 12 && (workSheet.Cells[linha, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == ""))
                    {
                        conteudo.Replace("[Código do Produto]", "[C.F.O.P. N.F. Entrada]");
                        conteudo.Replace("Cmp_Pro_id", "Cmp_CFOP");
                    }
                    else if (j == 9 && (workSheet.Cells[linha, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == ""))
                    {
                        conteudo.Replace("[Código do Produto]", "[Número da N.F.]");
                        conteudo.Replace("Cmp_Pro_id", "Cmp_NF_Entrada");
                    }
                    else if (j == 3 && (workSheet.Cells[linha, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == ""))
                    {
                        conteudo.Replace("[Código do Produto]", "[Código do Fornecedor]");
                        conteudo.Replace("Cmp_Pro_id", "Cmp_For_ID");
                    }
                    else if (j == 11 && (workSheet.Cells[linha, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == ""))
                    {
                        conteudo.Replace("[Código do Produto]", "[Data de Emissão NF]");
                        conteudo.Replace("Cmp_Pro_id", "Cmp_NF_DT");
                    }
                    else if (j == workSheet.Dimension.End.Column)
                    {

                        if (workSheet.Cells[linha, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == "")
                        {
                            conteudo.Append(" ', " + pegarID("D_Compras") + ") ");
                        }
                        else
                        {
                            conteudo.Append(" " + workSheet.Cells[linha, j].Value.ToString() + "', " + pegarID("D_Compras") + ") ");
                        }
                    }
                    else if (workSheet.Cells[linha, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == "")
                    {

                        conteudo.Append(" ");
                    }
                    else
                    {
                        conteudo.Append(" " + workSheet.Cells[linha, j].Value.ToString() + " ");
                    }
                }
                //  conn = new SqlConnection("Data Source=BRCAENRODRIGUES\\SQLEXPRESS01; Integrated Security=True; Initial Catalog=LAMPADA");
                conn.Open();
                SqlCommand cmda = conn.CreateCommand();
                cmda.CommandText = conteudo.ToString();
                SqlTransaction trEa = null;
                trEa = conn.BeginTransaction();
                cmda.Transaction = trEa;
                cmda.ExecuteNonQuery();
                trEa.Commit();
                conn.Close();
                conteudo.Clear();
                package.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public int pegarID(string tabela)
        {
            try
            {
                int arqId = 0;
                string oString = "select max(Arq_ID) from S_ArquivoCarregado where Arq_Tabela = @tabela";
                SqlCommand oCmd = new SqlCommand(oString, conn);
                oCmd.Parameters.AddWithValue("@tabela", tabela);

                if (conn.State.ToString() == "Closed")
                {
                    conn.Open();
                }
                using (SqlDataReader oReader = oCmd.ExecuteReader())
                {

                    while (oReader.Read())
                    {

                        if (oReader[0].ToString() == "")
                        {
                            arqId = 1;
                        }
                        else
                        {
                            arqId = Int16.Parse(oReader[0].ToString());
                            arqId = arqId + 1;
                        }
                    }

                }
                return arqId;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return 0;
            }
            finally
            {
                conn.Close();
            }
        }

        public void VendasPenLayout(int linha)
        {
            string filePath = caminho;
            try
            {

                ///// temporario
                //conn = new SqlConnection("Data Source=BRCAENRODRIGUES\\SQLEXPRESS01; Integrated Security=True; Initial Catalog=LAMPADA");
                //filePath = @"C:\Base\Vendas_Doosan_Jan_Jun_2019.xlsx";
                FileInfo existingFile = new FileInfo(filePath);
                ExcelPackage package = new ExcelPackage(existingFile);
                ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
                StringBuilder conteudo = new StringBuilder();
                SqlCommand cmd = conn.CreateCommand();

                for (int j = workSheet.Dimension.Start.Column; j <= workSheet.Dimension.End.Column; j++)
                {
                    if (j == 1)
                    {
                        conteudo.Append(" INSERT INTO[dbo].[A_PendenciaLayout] " +
                                            " ([pen_Linha], " +
                                            " [pen_Tabela]," +
                                            " [pen_Campo]," +
                                            " [pen_Posi]," +
                                            " [pen_Tam]," +
                                            " [pen_Erro]," +
                                            " [pen_Registro]," +
                                            " [pen_Arq_Origem])" +
                                            " VALUES ( ");
                        if (workSheet.Cells[linha, j].Value == null)
                        {
                            int registro = linha - 1;
                            conteudo.Append(" " + registro + ", 'D_Vendas_Itens', 'Vnd_Pro_id', 0, 0, 'Campo [Código Cliente] é obrigatório', '");
                        }
                        else
                        {
                            int registro = linha - 1;
                            conteudo.Append(" " + registro + ", 'D_Vendas_Itens', 'Vnd_Pro_id', 0, 0, 'Campo [Código Cliente] é obrigatório', '" + workSheet.Cells[linha, j].Value.ToString() + " ");
                        }
                    }
                    else if (j == 5 && (workSheet.Cells[linha, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == ""))
                    {
                        conteudo.Replace("[Código Cliente]", "[CFOP]");
                    }
                    else if (j == 2 && (workSheet.Cells[linha, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == ""))
                    {
                        conteudo.Replace("[Código Cliente]", "[Número NF]");
                    }
                    else if (j == 10 && (workSheet.Cells[linha, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == ""))
                    {
                        conteudo.Replace("[Código Cliente]", "[Código Produto]");
                    }
                    else if (j == 6 && (workSheet.Cells[linha, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == ""))
                    {
                        conteudo.Replace("[Código Cliente]", "[Data de Emissão NF]");
                    }
                    else if (j == workSheet.Dimension.End.Column)
                    {

                        if (workSheet.Cells[linha, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == "")
                        {
                            conteudo.Append(" ', " + pegarID("D_Vendas_Itens") + ") ");
                        }
                        else
                        {
                            conteudo.Append(" " + workSheet.Cells[linha, j].Value.ToString() + "', " + pegarID("D_Vendas_Itens") + ") ");
                        }
                    }
                    else if (workSheet.Cells[linha, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == "")
                    {

                        conteudo.Append(" ");
                    }
                    else
                    {
                        conteudo.Append(" " + workSheet.Cells[linha, j].Value.ToString() + " ");
                    }
                }

                //Clipboard.SetText(conteudo.ToString());
                conn.Open();
                cmd.CommandText = conteudo.ToString();
                SqlTransaction trE = null;
                trE = conn.BeginTransaction();
                cmd.Transaction = trE;
                cmd.ExecuteNonQuery();
                trE.Commit();
                conn.Close();
                conteudo.Clear();
                package.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Clientes()
        {
            int linha = 1;
            int numRepetidos = 0, numCarregados = 0, numPendencias = 0;
            FileInfo existingFile = new FileInfo(caminho);
            ExcelPackage package = new ExcelPackage(existingFile);
            ExcelWorksheet workSheet = package.Workbook.Worksheets[cmbPlanilha.SelectedIndex + 1];
            StringBuilder conteudo = new StringBuilder();
            var lista = new List<String>();
            SqlCommand cmd = conn.CreateCommand();
            ArrayList Excel = new ArrayList();
            ArrayList SQL = new ArrayList();
            ArrayList repetido = new ArrayList();
            ArrayList carregado = new ArrayList();
            bool pendencia = false;
            if (conn.State.ToString() == "Closed")
            {
                conn.Open();
            }

            SqlCommand cmdProc = conn.CreateCommand();
            SqlTransaction trProc = null;
            cmdProc.CommandText = "create or alter PROCEDURE [dbo].[SP_VERIFICA_CLIENTES_REPETIDOS_CARREGADOR] @CLI VARCHAR(MAX) AS BEGIN IF NOT EXISTS(SELECT * FROM D_Clientes WHERE Cli_Id = @CLI)  BEGIN  RETURN 0; END ELSE  RETURN 1;  END ";
            trProc = conn.BeginTransaction();
            cmdProc.Transaction = trProc;
            cmdProc.ExecuteNonQuery();
            trProc.Commit();

            for (int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
            {
                pendencia = false;

                for (int j = workSheet.Dimension.Start.Column; j <= workSheet.Dimension.End.Column; j++)
                {

                    if (j == 1 && (workSheet.Cells[i, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == ""))
                    {
                        pendencia = true;
                        numPendencias++;
                        lblPendencia.Text = numPendencias.ToString();
                    }

                    else if (j == 1)
                    {
                        Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US")
                        {
                            DateTimeFormat = { YearMonthPattern = "yyyy-mm-dd" }
                        };

                        SqlCommand cmdeProc = conn.CreateCommand();
                        cmdeProc.CommandType = CommandType.StoredProcedure;
                        cmdeProc.CommandText = "[SP_VERIFICA_CLIENTES_REPETIDOS_CARREGADOR]";
                        cmdeProc.Parameters.Add("@CLI", SqlDbType.VarChar);
                        cmdeProc.Parameters["@CLI"].Direction = ParameterDirection.ReturnValue;
                        cmdeProc.Parameters.AddWithValue("@CLI", workSheet.Cells[i, j].Value.ToString());

                        if (conn.State.ToString() == "Closed")
                            conn.Open();

                        cmdeProc.ExecuteNonQuery();
                        int ret = Convert.ToInt32(cmdeProc.Parameters["@CLI"].Value);

                        conn.Close();


                        if (ret == 1)
                        {
                            numRepetidos++;
                            lblRepetido.Text = numRepetidos.ToString();
                            lblRepetido.Refresh();
                        }
                        else
                        {
                            numCarregados++;
                            lblCarregada.Text = numCarregados.ToString();
                            lblCarregada.Refresh();
                        }

                        conteudo.Append(" declare @cli_id varchar(max) = '" + workSheet.Cells[i, j].Value.ToString() + "';");
                        conteudo.Append(Environment.NewLine);
                        conteudo.Append(" if  (select max(cli_id) from D_Clientes where cli_id = @cli_id) = (select (cli_id) from D_Clientes where cli_id = (@cli_id)) ");
                        conteudo.Append(Environment.NewLine);
                        conteudo.Append(" print 'OK' ");
                        conteudo.Append(Environment.NewLine);
                        conteudo.Append(" else ");
                        conteudo.Append(" INSERT INTO D_CLIENTES " +
                        "(CLI_ID, " +
                        "CLI_NOME, " +
                        "CLI_PSS_ID, " +
                        "CLI_VINC, " +
                        "CLI_VINC_DT_INI, " +
                        "CLI_VINC_DT_FIM, " +
                        "CLI_CNPJ, " +
                        "[Lin_Origem_ID], " +
                        "[Arq_Origem_ID]) " +
                        " VALUES ( ");
                        conteudo.Append("'" + workSheet.Cells[i, j].Value + "', ");
                    }
                    else if ((j == 5 || j == 6) && (workSheet.Cells[i, j].Value == null) || (workSheet.Cells[linha, j].Value.ToString() == ""))
                    {
                        conteudo.Append(workSheet.Cells[i, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == "" ? " NULL, " : " '" + workSheet.Cells[i, j].Value + "', ");
                    }
                    else if ((j == 5 || j == 6) && (workSheet.Cells[i, j].Value != null || (workSheet.Cells[i, j].Value.ToString() != "")))
                    {
                        if (workSheet.Cells[i, j].Value.ToString().Contains(@"/00/"))
                        {
                            conteudo.Append(" NULL, ");
                        }
                        else
                        {
                            conteudo.Append(" ' " + workSheet.Cells[i, j].Value + "', ");
                        }
                    }
                    else if ((workSheet.Cells[i, j].Value == null ? " NULL " : workSheet.Cells[i, j].Value.GetType().Name.ToString()) == "DateTime")
                    {
                        DateTime oDate = DateTime.ParseExact(workSheet.Cells[i, j].Value.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                        conteudo.Append("'" + oDate + "', ");
                    }

                    else if (j == workSheet.Dimension.End.Column)
                    {
                        conteudo.Append(workSheet.Cells[i, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == "" ? " NULL, '" + linha + "', " : " '" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "' , '" + linha + "', ");
                        conteudo.Append(" " + pegarID("D_Clientes") + " ");
                    }

                    else
                    {
                        conteudo.Append(workSheet.Cells[i, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == "" ? " NULL, " : " '" + workSheet.Cells[i, j].Value + "', ");
                    }
                }


                if (i == workSheet.Dimension.End.Row)
                {
                    conteudo.Append(" ) ");
                }
                else
                {
                    conteudo.Append(")");
                    conteudo.Append(Environment.NewLine);
                }

                //Clipboard.SetText(conteudo.ToString());
                if (conn.State.ToString() == "Closed")
                {
                    conn.Open();
                }

                if (pendencia == false)
                {
                    // Clipboard.SetText(conteudo.ToString());
                    cmd.CommandText = conteudo.ToString();
                    fazTransacao(conn, cmd);
                }

                conteudo.Clear();

            }
            package.Dispose();

            if (numCarregados == 0)
            {
                MessageBox.Show(new Form { TopMost = true }, "Nenhum registro de clientes carregado");
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Carregamento de " + numCarregados.ToString() + " registros de clientes realizados com sucesso");

                SqlCommand cmdArquivoCarregado = conn.CreateCommand();
                cmdArquivoCarregado.CommandText = gravaId(caminho, numCarregados, "D_Clientes");
                fazTransacao(conn, cmdArquivoCarregado);
                gravaId(caminho, numCarregados, "D_Clientes");
            }


            conteudo.Clear();


        }


        public void Fornecedores()
        {
            int linha = 1;
            int numRepetidos = 0, numCarregados = 0, numPendencias = 0;
            FileInfo existingFile = new FileInfo(caminho);
            ExcelPackage package = new ExcelPackage(existingFile);
            ExcelWorksheet workSheet = package.Workbook.Worksheets[cmbPlanilha.SelectedIndex + 1];
            StringBuilder conteudo = new StringBuilder();
            var lista = new List<String>();
            SqlCommand cmd = conn.CreateCommand();
            ArrayList Excel = new ArrayList();
            ArrayList SQL = new ArrayList();
            ArrayList repetido = new ArrayList();
            ArrayList carregado = new ArrayList();
            bool pendencia = false;

            if (conn.State.ToString() == "Closed")
            {
                conn.Open();
            }

            SqlCommand cmdProc = conn.CreateCommand();
            SqlTransaction trProc = null;
            cmdProc.CommandText = "create or alter PROCEDURE [dbo].[SP_VERIFICA_FORNECEDORES_REPETIDOS_CARREGADOR] @FOR VARCHAR(MAX) AS BEGIN IF NOT EXISTS(SELECT * FROM D_Fornecedores WHERE For_Id = @FOR)  BEGIN  RETURN 0; END ELSE  RETURN 1;  END ";
            trProc = conn.BeginTransaction();
            cmdProc.Transaction = trProc;
            cmdProc.ExecuteNonQuery();
            trProc.Commit();

            for (int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
            {
                pendencia = false;

                for (int j = workSheet.Dimension.Start.Column; j <= workSheet.Dimension.End.Column; j++)
                {
                    if (j == 1 && (workSheet.Cells[i, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == ""))
                    {
                        pendencia = true;
                        numPendencias++;
                        lblPendencia.Text = numPendencias.ToString();
                    }
                    else if (j == workSheet.Dimension.End.Column)
                    {
                        conteudo.Append(workSheet.Cells[i, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == "" ? " NULL, '" + linha + "', " : " '" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "' , '" + linha + "', ");
                        conteudo.Append(" " + pegarID("D_Fornecedores") + " ");
                    }
                    else if (j == 1)
                    {
                        Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US")
                        {
                            DateTimeFormat = { YearMonthPattern = "yyyy-mm-dd" }
                        };

                        SqlCommand cmdeProc = conn.CreateCommand();
                        cmdeProc.CommandType = CommandType.StoredProcedure;
                        cmdeProc.CommandText = "[SP_VERIFICA_FORNECEDORES_REPETIDOS_CARREGADOR]";
                        cmdeProc.Parameters.Add("@FOR", SqlDbType.VarChar);
                        cmdeProc.Parameters["@FOR"].Direction = ParameterDirection.ReturnValue;
                        cmdeProc.Parameters.AddWithValue("@FOR", workSheet.Cells[i, j].Value.ToString());

                        if (conn.State.ToString() == "Closed")
                            conn.Open();

                        cmdeProc.ExecuteNonQuery();
                        int ret = Convert.ToInt32(cmdeProc.Parameters["@FOR"].Value);

                        conn.Close();

                        if (ret == 1)
                        {
                            numRepetidos++;
                            lblRepetido.Text = numRepetidos.ToString();
                            lblRepetido.Refresh();
                        }
                        else
                        {
                            numCarregados++;
                            lblCarregada.Text = numCarregados.ToString();
                            lblCarregada.Refresh();
                        }

                        if (j == 1)
                        {
                            conteudo.Append(" declare @for_id varchar(max) = '" + workSheet.Cells[i, j].Value + "';");
                            conteudo.Append(Environment.NewLine);
                            conteudo.Append(" if  (select max(for_id) from D_Fornecedores where for_id = @for_id) = (select (for_id) from D_Fornecedores where for_id = (@for_id)) ");
                            conteudo.Append(Environment.NewLine);
                            conteudo.Append(" print 'OK' ");
                            conteudo.Append(Environment.NewLine);
                            conteudo.Append(" else ");
                            conteudo.Append(" INSERT INTO D_Fornecedores " +
                            " (for_ID, " +
                            " for_NOME, " +
                            " for_PSS_ID, " +
                            " for_VINC, " +
                            " for_VINC_DT_INI, " +
                            " for_VINC_DT_FIM, " +
                            " for_CNPJ, " +
                            " [Lin_Origem_ID], " +
                            " [Arq_Origem_ID]) " +
                            " VALUES ( ");
                            conteudo.Append("'" + workSheet.Cells[i, j].Value + "', ");
                        }
                    }
                    else if ((j == 5 || j == 6) && (workSheet.Cells[i, j].Value == null) || (workSheet.Cells[linha, j].Value.ToString() == ""))
                    {
                        conteudo.Append(workSheet.Cells[i, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == "" ? " NULL, " : " '" + workSheet.Cells[i, j].Value + "', ");
                    }
                    else if ((j == 5 || j == 6) && (workSheet.Cells[i, j].Value != null || (workSheet.Cells[i, j].Value.ToString() != "")))
                    {
                        if (workSheet.Cells[i, j].Value.ToString().Contains(@"/00/"))
                        {
                            conteudo.Append(" NULL, ");
                        }
                        else
                        {
                            conteudo.Append(" ' " + workSheet.Cells[i, j].Value.ToString().Replace("'", "" + "', "));
                        }
                    }
                    else if ((workSheet.Cells[i, j].Value == null ? " NULL " : workSheet.Cells[i, j].Value.GetType().Name.ToString()) == "DateTime")
                    {
                        DateTime oDate = DateTime.ParseExact(workSheet.Cells[i, j].Value.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                        conteudo.Append("'" + oDate + "', ");
                    }
                    else
                    {
                        conteudo.Append(" '" + workSheet.Cells[i, j].Value.ToString().Replace("'", "") + "', ");
                    }

                }

                if (i == workSheet.Dimension.End.Row)
                {
                    conteudo.Append(" ) ");
                }
                else
                {
                    conteudo.Append(")");
                    conteudo.Append(Environment.NewLine);
                }

                //Clipboard.SetText(conteudo.ToString());
                if (conn.State.ToString() == "Closed")
                {
                    conn.Open();
                }

                if (pendencia == false)
                {
                    //  Clipboard.SetText(conteudo.ToString());
                    cmd.CommandText = conteudo.ToString();
                    fazTransacao(conn, cmd);
                }

                conteudo.Clear();
            }
            package.Dispose();

            if (numCarregados == 0)
            {
                MessageBox.Show(new Form { TopMost = true }, "Nenhum registro de fornecedores carregado");
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Carregamento de " + numCarregados.ToString() + " registros de fornecedores realizados com sucesso");

                SqlCommand cmdArquivoCarregado = conn.CreateCommand();
                cmdArquivoCarregado.CommandText = gravaId(caminho, numCarregados, "D_Fornecedores");
                fazTransacao(conn, cmdArquivoCarregado);
                gravaId(caminho, numCarregados, "D_Fornecedores");
            }
            conteudo.Clear();
        }


        public void Inventario()
        {

                int linha = 1;
                string filePath = caminho;
                int numRepetidos = 0, numCarregados = 0, numPendencias = 0;
                //try
                //{
                //conn = new SqlConnection("Data Source=BRCAENRODRIGUES\\SQLEXPRESS01; Integrated Security=True; Initial Catalog=LAMPADA");
                //string filePath = @"C:\Base\saldos_maio_evonik_2019.xlsx";
                // ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
                FileInfo existingFile = new FileInfo(filePath);
                ExcelPackage package = new ExcelPackage(existingFile);
                ExcelWorksheet workSheet = package.Workbook.Worksheets[cmbPlanilha.SelectedIndex + 1];
                StringBuilder conteudo = new StringBuilder();
                var lista = new List<String>();
                SqlCommand cmd = conn.CreateCommand();
                string produto = "", cnpj = "";
                bool pendencia = false;
             
                if (conn.State.ToString() == "Closed")
                {
                    conn.Open();
                }

                SqlCommand cmdProc = conn.CreateCommand();
                SqlTransaction trProc = null;
                cmdProc.CommandText = "CREATE or ALTER PROCEDURE [dbo].[SP_VERIFICA_INVENTARIO_REPETIDOS_CARREGADOR] @PRO_ID VARCHAR(MAX), @CNPJ VARCHAR(MAX)  AS BEGIN IF NOT EXISTS(SELECT * FROM D_Inventario_Carga WHERE inv_pro_id = @PRO_ID and Inv_CNPJ = @CNPJ)  BEGIN  RETURN 0; END ELSE  RETURN 1;  END ";
                trProc = conn.BeginTransaction();
                cmdProc.Transaction = trProc;
                cmdProc.ExecuteNonQuery();
                trProc.Commit();
             
                for (int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
                {
                    pendencia = false;

                   for (int j = workSheet.Dimension.Start.Column; j <= workSheet.Dimension.End.Column; j++)
                    {
                        if (j == 1 && (workSheet.Cells[i, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == ""))
                        {
                        pendencia = true;
                        numPendencias++;
                        lblPendencia.Text = numPendencias.ToString();
                        }

                    if (j == workSheet.Dimension.End.Column)
                        {
                            cnpj = (workSheet.Cells[i, j].Value == null ? " " : workSheet.Cells[i, j].Value.ToString()); ;
                            conteudo.Append(workSheet.Cells[i, j].Value == null ? " '' , '" + linha + "', " : " '" + workSheet.Cells[i, j].Value.ToString() + "' , '" + linha + "', ");
                            conteudo.Append(" " + pegarID("D_INVENTARIO_CARGA") + "  ");
                        }
                        else if ((j == 1) && workSheet.Cells[i, j].Value != null)
                        {

                        SqlCommand cmdeProc = conn.CreateCommand();
                        cmdeProc.CommandType = CommandType.StoredProcedure;
                        cmdeProc.CommandText = "[SP_VERIFICA_INVENTARIO_REPETIDOS_CARREGADOR]";
                        cmdeProc.Parameters.Add("@PRO_ID", SqlDbType.VarChar);
                        cmdeProc.Parameters["@PRO_ID"].Direction = ParameterDirection.ReturnValue;
                        cmdeProc.Parameters.AddWithValue("@PRO_ID", produto);
                        cmdeProc.Parameters.Add("@CNPJ", SqlDbType.VarChar);
                        cmdeProc.Parameters["@CNPJ"].Direction = ParameterDirection.ReturnValue;
                        cmdeProc.Parameters.AddWithValue("@CNPJ", cnpj);

                        if (conn.State.ToString() == "Closed")
                            conn.Open();

                        cmdeProc.ExecuteNonQuery();
                        int ret = Convert.ToInt32(cmdeProc.Parameters["@PRO_ID"].Value);

                        conn.Close();

                        if (ret == 1)
                        {
                            numRepetidos++;
                            lblRepetido.Text = numRepetidos.ToString();
                            lblRepetido.Refresh();
                        }
                        else
                        {
                            numCarregados++;
                            lblCarregada.Text = numCarregados.ToString();
                            lblCarregada.Refresh();
                        }

                        produto = workSheet.Cells[i, j].Value.ToString();

                            conteudo.Append(" declare @inv_pro_id varchar(max)  = '" + produto + "';");
                            conteudo.Append(Environment.NewLine);
                            conteudo.Append(" declare @cnpj varchar(max) = '" + cnpj + "';");
                           // MessageBox.Show("Test " + cnpj);
                            conteudo.Append(Environment.NewLine);
                            conteudo.Append("  if  (select max(inv_pro_id) from D_Inventario_Carga where inv_pro_id = @inv_pro_id and  Inv_CNPJ = @cnpj) > '' ");
                            conteudo.Append(Environment.NewLine);
                            conteudo.Append(" print 'OK' ");
                            conteudo.Append(Environment.NewLine);
                            conteudo.Append(" else ");
                            conteudo.Append(" INSERT INTO D_INVENTARIO_CARGA " +
                                            " (INV_PRO_ID, " +
                                            " INV_DATA, " +
                                            " INV_QTDE, " +
                                            " INV_VALOR, " +
                                            " INV_UND_ID, " +
                                            " INV_CNPJ, " +
                                            " [Lin_Origem_ID], " +
                                            " [Arq_Origem_ID]) " +
                                            " VALUES ( ");
                            conteudo.Append(" '" + workSheet.Cells[i, j].Value.ToString() + "', ");
                        }
  
                        else
                        {
                         
                            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US")
                            {
                                DateTimeFormat = { YearMonthPattern = "yyyy-mm-dd" }
                            };
                            if ((j == 2) && workSheet.Cells[i, j].Value != null && workSheet.Cells[i, j].Value.GetType().Name.ToString() != "DateTime")
                            {
                                DateTime oDate = DateTime.ParseExact(workSheet.Cells[i, j].Value.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                                conteudo.Append("'" + oDate + "', ");
                            }
                            else
                            {
                                if ((workSheet.Cells[i, j].Value == null ? " NULL " : workSheet.Cells[i, j].Value.GetType().Name.ToString()) == "DateTime")
                                {
                                    //DateTime oDate = DateTime.ParseExact(workSheet.Cells[i, j].Value.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                                    //conteudo.Append("'" + oDate + "', ");
                                    conteudo.Append("'" + workSheet.Cells[i, j].Value.ToString() + "', ");
                                }
                                else
                                {
                                    conteudo.Append(workSheet.Cells[i, j].Value == null ? " NULL, " : "'" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "', ");
                                }
                            }
                        }
                    }
                    if (i == workSheet.Dimension.End.Row)
                    {
                        conteudo.Append(" ) ");
                    }
                    else
                    {
                        conteudo.Append(")");
                        conteudo.Append(Environment.NewLine);
                    }
                    conn.Open();
                    Clipboard.SetText(conteudo.ToString());
                    linha = linha + 1;
                    cmd.CommandText = conteudo.ToString();
                    SqlTransaction trE = null;
                    trE = conn.BeginTransaction();
                    cmd.Transaction = trE;
                    cmd.ExecuteNonQuery();
                    trE.Commit();
                    conteudo.Clear();
                }
                package.Dispose();

         //       SqlCommand cmdArquivoCarregado = conn.CreateCommand();
                //cmdArquivoCarregado.CommandText =
                //" declare @tabela varchar(max) = 'D_Inventario_Carga';" +
                //" if (select count(arq_id) from S_ArquivoCarregado where Arq_Tabela = @tabela) = 0" +
                //" insert into S_ArquivoCarregado" +
                //" (Arq_ID, Arq_Nome, Arq_Tabela, Arq_Mensagem, Arq_DataCarga, Arq_Quantidade, Arq_Login)" +
                //" values(1, '" + caminho + "', @tabela, 'Carga efetuada com sucesso.'," +
                //" GETDATE(), ' " + lblCarregada.ToString() + "', REPLACE(SUSER_NAME(), 'ATRAME\\',''))" +
                //" else" +
                //" insert into S_ArquivoCarregado" +
                //" (Arq_ID, Arq_Nome, Arq_Tabela, Arq_Mensagem, Arq_DataCarga, Arq_Quantidade, Arq_Login)" +
                //" values(" + pegarID("D_Inventario_Carga") + ", '" + caminho + "', @tabela, 'Carga efetuada com sucesso.'," +
                //" GETDATE(), " + lblCarregada.ToString() + ", REPLACE(SUSER_NAME(), 'ATRAME\\',''))";
             
            if (numCarregados == 0)
            {
                MessageBox.Show(new Form { TopMost = true }, "Nenhum registro de fornecedores carregado");
            }
            else
            {
                MessageBox.Show(new Form { TopMost = true }, "Carregamento de " + numCarregados.ToString() + " registros de inventario realizados com sucesso");

                SqlCommand cmdArquivoCarregado = conn.CreateCommand();
                cmdArquivoCarregado.CommandText = gravaId(caminho, numCarregados, "D_Inventario_Carga");
                fazTransacao(conn, cmdArquivoCarregado);
                gravaId(caminho, numCarregados, "D_Inventario_Carga");
            }
            conteudo.Clear();

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
            //finally
            //{
            //    MessageBox.Show(new Form { TopMost = true }, "Carregamento de " + linha + " registros de inventario realizados com sucesso");
            //}
        }


        public void Insumo_Produto()
        {
            string filePath = caminho;
            //conn = new SqlConnection("Data Source=BRCAENRODRIGUES\\SQLEXPRESS01; Integrated Security=True; Initial Catalog=LAMPADA");
            //string filePath = @"C:\Base\ordens_de_producao_evonik_2019.xlsx";
            //ExcelWorksheet workSheet = package.Workbook.Worksheets.First();

            FileInfo existingFile = new FileInfo(filePath);
            ExcelPackage package = new ExcelPackage(existingFile);
            ExcelWorksheet workSheet = package.Workbook.Worksheets[cmbPlanilha.SelectedIndex + 1];
            StringBuilder conteudo = new StringBuilder();
            var lista = new List<String>();
            SqlCommand cmd = conn.CreateCommand();
            int linha = 1;
            if(conn.State.ToString() =="Closed")
            conn.Open();

            lblTotal.Text = workSheet.Dimension.End.Row.ToString();
            lblTotal.Refresh();

            try
            {

                for (int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
                {
                    lblCarregada.Text = i.ToString();
                    lblCarregada.Refresh();

                    for (int j = workSheet.Dimension.Start.Column; j <= workSheet.Dimension.End.Column; j++)
                    {
                        if (j == workSheet.Dimension.End.Column)
                        {
                            conteudo.Append(workSheet.Cells[i, j].Value == null ? "0, '" + linha + "', " : " '" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "' , '" + linha + "', ");
                            conteudo.Append(" " + pegarID("D_Insumo_Produto") + " ");
                        }
                        else
                        {
                            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US")
                            {
                                DateTimeFormat = { YearMonthPattern = "yyyy-mm-dd" }
                            };

                            if (j == 1)
                            {
                                conteudo.Append(" INSERT INTO D_INSUMO_PRODUTO " +
                                     "([Ins_PA_Pro_Id] " +
                                      ",[Ins_Qtd_Produzida] " +
                                      ",[Ins_PA_Und_Id] " +
                                      ",[Ins_MP_Pro_Id] " +
                                      ",[Ins_Qtd_Requisitada] " +
                                      ",[Ins_MP_Und_Id] " +
                                      ",[Ins_Ordem_Prod] " +
                                      ",[Ins_DT_Ini] " +
                                      ",[Ins_DT_Fim] " +
                                      ",[Ins_CNPJ] " +
                                      ",[Lin_Origem_ID] " +
                                      ",[Arq_Origem_ID]) " +
                                " VALUES ( ");
                                conteudo.Append("'" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "', ");
                            }
                            else if ((j == 8 || j == 9) && workSheet.Cells[i, j].Value != null && workSheet.Cells[i, j].Value.GetType().Name.ToString() != "DateTime")
                            {
                                DateTime dt = DateTime.FromOADate(Convert.ToInt64(workSheet.Cells[i, j].Value));
                                conteudo.Append("'" + dt + "', ");
                            }
                            else if ((j == 3 || j == 6) && workSheet.Cells[i, j].Value != null)
                            {
                                Unidades unidade = (Unidades)System.Enum.Parse(typeof(Unidades), workSheet.Cells[i, j].Value.ToString());
                                conteudo.Append("'" + ((int)unidade).ToString() + "', ");
                            }
                            else
                            {
                                if ((workSheet.Cells[i, j].Value == null ? " NULL " : workSheet.Cells[i, j].Value.GetType().Name.ToString()) == "DateTime")
                                {
                                    workSheet.Cells[i, j].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.YearMonthPattern;
                                    conteudo.Append(workSheet.Cells[i, j].Value == null ? " NULL, " : "'" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "', ");
                                }
                                else
                                {
                                    conteudo.Append(workSheet.Cells[i, j].Value == null ? " NULL, " : "'" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "', ");
                                }
                            }
                        }
                    }
                    if (i == workSheet.Dimension.End.Row)
                    {
                        conteudo.Append(" ) ");
                    }
                    else
                    {
                        conteudo.Append(")");
                        conteudo.Append(Environment.NewLine);
                    }
                    conn.Open();
                    //Clipboard.SetText(conteudo.ToString());
                    linha = linha + 1;
                    cmd.CommandText = conteudo.ToString();
                    SqlTransaction trE = null;
                    trE = conn.BeginTransaction();
                    cmd.Transaction = trE;
                    cmd.ExecuteNonQuery();
                    trE.Commit();
                    conteudo.Clear();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                package.Dispose();
                MessageBox.Show(" Carregamento de " + linha.ToString() + " registros de Insunmo Produto/Ordens de Produção");
                SqlCommand cmdArquivoCarregado = conn.CreateCommand();
                cmdArquivoCarregado.CommandText =
                " declare @tabela varchar(max) = 'D_Insumo_Produto';" +
                " if (select count(arq_id) from S_ArquivoCarregado where Arq_Tabela = @tabela) = 0" +
                " insert into S_ArquivoCarregado" +
                " (Arq_ID, Arq_Nome, Arq_Tabela, Arq_Mensagem, Arq_DataCarga, Arq_Quantidade, Arq_Login)" +
                " values(1, '" + caminho + "', @tabela, 'Carga efetuada com sucesso.'," +
                " GETDATE(), ' " + linha.ToString() + "', REPLACE(SUSER_NAME(), 'ATRAME\\',''))" +
                " else" +
                " insert into S_ArquivoCarregado" +
                " (Arq_ID, Arq_Nome, Arq_Tabela, Arq_Mensagem, Arq_DataCarga, Arq_Quantidade, Arq_Login)" +
                " values(" + pegarID("D_Insumo_Produto") + ", '" + caminho + "', @tabela, 'Carga efetuada com sucesso.'," +
                " GETDATE(), " + linha.ToString() + ", REPLACE(SUSER_NAME(), 'ATRAME\\',''))";
                conn.Open();
                SqlTransaction trA = null;
                trA = conn.BeginTransaction();
                cmdArquivoCarregado.Transaction = trA;
                cmdArquivoCarregado.ExecuteNonQuery();
                trA.Commit();
                conn.Close();
            }
        }

        public void Relacao_Carga()
        {
            string filePath = caminho;

            //conn = new SqlConnection("Data Source=BRCAENRODRIGUES\\SQLEXPRESS01; Integrated Security=True; Initial Catalog=LAMPADA");
            //string filePath = @"C:\Base\ordens_de_producao_evonik_2019.xlsx";

            FileInfo existingFile = new FileInfo(filePath);
            ExcelPackage package = new ExcelPackage(existingFile);
            ExcelWorksheet workSheet = package.Workbook.Worksheets[cmbPlanilha.SelectedIndex + 1];
            //ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
            StringBuilder conteudo = new StringBuilder();
            var lista = new List<String>();
            SqlCommand cmd = conn.CreateCommand();
            int linha = 1;

            lblTotal.Text = workSheet.Dimension.End.Row.ToString();
            lblTotal.Refresh();

            try
            {
                for (int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
                {
                    lblCarregada.Text = i.ToString();
                    lblCarregada.Refresh();

                    for (int j = workSheet.Dimension.Start.Column; j <= workSheet.Dimension.End.Column; j++)
                    {
                        if (j == workSheet.Dimension.End.Column)
                        {
                            conteudo.Append(workSheet.Cells[i, j].Value == null ? " NULL, '" + linha + "', " : " '" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "' , '" + linha + "', ");
                            conteudo.Append(" " + pegarID("D_Relacao_Carga") + " ");
                        }
                        else
                        {
                            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US")
                            {
                                DateTimeFormat = { YearMonthPattern = "yyyy-mm-dd" }
                            };

                            if (j == 1)
                            {
                                conteudo.Append(" INSERT INTO D_RELACAO_CARGA " +
                                                " ([REL_PA_PRO_ID] " +
                                                " ,[REL_MP_PRO_ID] " +
                                                " ,[REL_PRODUZIDA] " +
                                                " ,[REL_REQUISITADA] " +
                                                " ,[REL_RELACAO] " +
                                                " ,[REL_TIPO] " +
                                                " ,[Lin_Origem_ID] " +
                                                " ,[Arq_Origem_ID]) " +
                                                " VALUES ( ");
                                conteudo.Append("'" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "', ");
                            }
                            else if (workSheet.Cells[i, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == "")
                            {
                                conteudo.Append(" NULL, ");
                            }
                            else
                            {
                                conteudo.Append(workSheet.Cells[i, j].Value == null ? " NULL, " : "'" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "', ");
                            }
                        }
                    }

                    if (i == workSheet.Dimension.End.Row)
                    {
                        conteudo.Append(" ) ");
                    }
                    else
                    {
                        conteudo.Append(")");
                        conteudo.Append(Environment.NewLine);
                    }
                    conn.Open();
                    //Clipboard.SetText(conteudo.ToString());
                    linha = linha + 1;
                    cmd.CommandText = conteudo.ToString();
                    SqlTransaction trE = null;
                    trE = conn.BeginTransaction();
                    cmd.Transaction = trE;
                    cmd.ExecuteNonQuery();
                    trE.Commit();
                    conteudo.Clear();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                package.Dispose();
                MessageBox.Show(new Form { TopMost = true }, "Carregamento " + linha.ToString() + " registros de Relação Produção de Produção realizado com sucesso!");

                SqlCommand cmdArquivoCarregado = conn.CreateCommand();
                cmdArquivoCarregado.CommandText =
                " declare @tabela varchar(max) = 'D_RELACAO_CARGA';" +
                " if (select count(arq_id) from S_ArquivoCarregado where Arq_Tabela = @tabela) = 0" +
                " insert into S_ArquivoCarregado" +
                " (Arq_ID, Arq_Nome, Arq_Tabela, Arq_Mensagem, Arq_DataCarga, Arq_Quantidade, Arq_Login)" +
                " values(1, '" + caminho + "', @tabela, 'Carga efetuada com sucesso.'," +
                " GETDATE(), ' " + linha.ToString() + "', REPLACE(SUSER_NAME(), 'ATRAME\\',''))" +
                " else" +
                " insert into S_ArquivoCarregado" +
                " (Arq_ID, Arq_Nome, Arq_Tabela, Arq_Mensagem, Arq_DataCarga, Arq_Quantidade, Arq_Login)" +
                " values(" + pegarID("D_RELACAO_CARGA") + ", '" + caminho + "', @tabela, 'Carga efetuada com sucesso.'," +
                " GETDATE(), " + linha.ToString() + ", REPLACE(SUSER_NAME(), 'ATRAME\\',''))";

                conn.Open();
                SqlTransaction trA = null;
                trA = conn.BeginTransaction();
                cmdArquivoCarregado.Transaction = trA;
                cmdArquivoCarregado.ExecuteNonQuery();
                trA.Commit();
                conn.Close();
            }
        }

        public void PIC()
        {
            string filePath = caminho;

            //conn = new SqlConnection("Data Source=BRCAENRODRIGUES\\SQLEXPRESS01; Integrated Security=True; Initial Catalog=LAMPADA");
            //string filePath = @"C:\Base\ordens_de_producao_evonik_2019.xlsx";

            FileInfo existingFile = new FileInfo(filePath);
            ExcelPackage package = new ExcelPackage(existingFile);
            ExcelWorksheet workSheet = package.Workbook.Worksheets[cmbPlanilha.SelectedIndex + 1];
            //ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
            StringBuilder conteudo = new StringBuilder();
            var lista = new List<String>();
            SqlCommand cmd = conn.CreateCommand();
            int linha = 1;

            lblTotal.Text = workSheet.Dimension.End.Row.ToString();
            lblTotal.Refresh();

            try
            {
                for (int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
                {
                    lblCarregada.Text = i.ToString();
                    lblCarregada.Refresh();

                    for (int j = workSheet.Dimension.Start.Column; j <= workSheet.Dimension.End.Column; j++)
                    {
                        if (j == workSheet.Dimension.End.Column)
                        {
                            conteudo.Append(workSheet.Cells[i, j].Value == null ? " NULL, '" + linha + "', " : " '" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "' , '" + linha + "', ");
                            conteudo.Append(" " + pegarID("D_PIC") + " ");
                        }
                        else
                        {
                            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US")
                            {
                                DateTimeFormat = { YearMonthPattern = "yyyy-mm-dd" }
                            };

                            if (j == 1)
                            {
                                conteudo.Append(" INSERT INTO D_PIC " +
                                "([Pic_For_Id] " +
                                ",[PIC_For_Pais] " +
                                ",[Pic_Pro_Id] " +
                                ",[Pic_Qtde] " +
                                ",[Pic_Vl_Moeda_Bruto] " +
                                ",[Pic_Cod_Moeda] " +
                                ",[Pic_Dt_Emissao] " +
                                ",[Pic_Dt_Venc] " +
                                ",[Pic_Cli_Id] " +
                                ",[PIC_Cli_Pais] " +
                                ",[Pic_Doc_Oper] " +
                                ",[Pic_Dias] " +
                                ",[Pic_CFOP] " +
                                ",[Pic_Imposto_Intern] " +
                                ",[Pic_FOB_Moeda] " +
                                ",[Pic_IPI] " +
                                ",[Pic_ICMS] " +
                                ",[Pic_Pis] " +
                                ",[Pic_Cofins] " +
                                ",[Pic_Frete] " +
                                ",[Pic_Seguro] " +
                                ",[Lin_Origem_ID] " +
                                ",[Arq_Origem_ID]) " +
                                " VALUES ( ");
                                conteudo.Append("'" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "', ");
                            }
                            else
                            {
                                if ((workSheet.Cells[i, j].Value == null ? " NULL " : workSheet.Cells[i, j].Value.GetType().Name.ToString()) == "DateTime")
                                {
                                    workSheet.Cells[i, j].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.YearMonthPattern;
                                    conteudo.Append(workSheet.Cells[i, j].Value == null ? " NULL, " : "'" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "', ");
                                }
                                else if ((j == 14 || j == 15 || j == 16 || j == 17 || j == 18 || j == 19 || j == 20 || j == 21) && ((workSheet.Cells[i, j].Value == null) ||workSheet.Cells[linha, j].Value.ToString() == ""))
                                {
                                    conteudo.Append(" " + 0 + ", ");
                                  
                                }
                                else if (workSheet.Cells[i, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == "")
                                {
                                    conteudo.Append(" NULL, ");
                                    
                                }
                                else
                                {
                                    conteudo.Append("'" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "', ");
                                  
                                }
                            }
                        }
                    }

                    if (i == workSheet.Dimension.End.Row)
                    {
                        conteudo.Append(" ) ");
                    }
                    else
                    {
                        conteudo.Append(")");
                        conteudo.Append(Environment.NewLine);
                    }
                    conn.Open();
                    Clipboard.SetText(conteudo.ToString());
                    linha = linha + 1;
                    cmd.CommandText = conteudo.ToString();
                    SqlTransaction trE = null;
                    trE = conn.BeginTransaction();
                    cmd.Transaction = trE;
                    cmd.ExecuteNonQuery();
                    trE.Commit();
                    conteudo.Clear();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                package.Dispose();
                MessageBox.Show(new Form { TopMost = true }, "Carregamento " + linha.ToString() + " registros de PIC realizado com sucesso!");

                SqlCommand cmdArquivoCarregado = conn.CreateCommand();
                cmdArquivoCarregado.CommandText =
                " declare @tabela varchar(max) = 'D_PIC';" +
                " if (select count(arq_id) from S_ArquivoCarregado where Arq_Tabela = @tabela) = 0" +
                " insert into S_ArquivoCarregado" +
                " (Arq_ID, Arq_Nome, Arq_Tabela, Arq_Mensagem, Arq_DataCarga, Arq_Quantidade, Arq_Login)" +
                " values(1, '" + caminho + "', @tabela, 'Carga efetuada com sucesso.'," +
                " GETDATE(), ' " + linha.ToString() + "', REPLACE(SUSER_NAME(), 'ATRAME\\',''))" +
                " else" +
                " insert into S_ArquivoCarregado" +
                " (Arq_ID, Arq_Nome, Arq_Tabela, Arq_Mensagem, Arq_DataCarga, Arq_Quantidade, Arq_Login)" +
                " values(" + pegarID("D_PIC") + ", '" + caminho + "', @tabela, 'Carga efetuada com sucesso.'," +
                " GETDATE(), " + linha.ToString() + ", REPLACE(SUSER_NAME(), 'ATRAME\\',''))";

                conn.Open();
                SqlTransaction trA = null;
                trA = conn.BeginTransaction();
                cmdArquivoCarregado.Transaction = trA;
                cmdArquivoCarregado.ExecuteNonQuery();
                trA.Commit();
                conn.Close();
            }
        }

        public void Custo()
        {
            string filePath = caminho;
            //conn = new SqlConnection("Data Source=BRCAENRODRIGUES\\SQLEXPRESS01; Integrated Security=True; Initial Catalog=LAMPADA");
            //string filePath = @"C:\Base\custo.xlsx";
            FileInfo existingFile = new FileInfo(filePath);
            ExcelPackage package = new ExcelPackage(existingFile);
            ExcelWorksheet workSheet = package.Workbook.Worksheets[cmbPlanilha.SelectedIndex + 1];
            //ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
            StringBuilder conteudo = new StringBuilder();
            var lista = new List<String>();
            SqlCommand cmd = conn.CreateCommand();
            int linha = 1;

            lblTotal.Text = workSheet.Dimension.End.Row.ToString();
            lblTotal.Refresh();

            try
            {
                for (int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
                {
                    lblCarregada.Text = i.ToString();
                    lblCarregada.Refresh();

                    for (int j = workSheet.Dimension.Start.Column; j <= workSheet.Dimension.End.Column; j++)
                    {
                        if ((j == 1))
                        {

                            conteudo.Append(" INSERT INTO D_Custo_Medio " +
                                "(Cst_Pro_Id, " +
                                "Cst_Mes, " +
                                "Cst_Ano, " +
                                "Cst_Vl_Custo, " +
                                "Cst_CNPJ, " +
                                "[Lin_Origem_ID], " +
                                "[Arq_Origem_ID]) " +
                                " VALUES ( ");
                        }
                        if (j == workSheet.Dimension.End.Column)
                        {
                            conteudo.Append(workSheet.Cells[i, j].Value == null ? "0, '" + linha + "', " : " '" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "' , '" + linha + "', ");
                            conteudo.Append(" " + pegarID("D_Custo_Medio") + " ");
                        }
                        else
                        {
                            if ((workSheet.Cells[i, j].Value == null ? " NULL " : workSheet.Cells[i, j].Value.GetType().Name.ToString()) == "DateTime")
                            {
                                workSheet.Cells[i, j].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.YearMonthPattern;
                                conteudo.Append(workSheet.Cells[i, j].Value == null ? " NULL, " : "'" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "', ");
                            }
                            else
                            {
                                conteudo.Append(workSheet.Cells[i, j].Value == null ? " NULL, " : "'" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "', ");
                            }
                        }
                    }
                    if (i == workSheet.Dimension.End.Row)
                    {
                        conteudo.Append(" ) ");
                    }
                    else
                    {
                        conteudo.Append(")");
                        conteudo.Append(Environment.NewLine);
                    }
                    //Clipboard.SetText(conteudo.ToString());
                    conn.Open();
                    linha = linha + 1;
                    cmd.CommandText = conteudo.ToString();
                    SqlTransaction trE = null;
                    trE = conn.BeginTransaction();
                    cmd.Transaction = trE;
                    cmd.ExecuteNonQuery();
                    trE.Commit();
                    conteudo.Clear();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                package.Dispose();
                MessageBox.Show(new Form { TopMost = true }, "Carregamento de " + linha.ToString() + " registros de  Custo Médio");
            }

            SqlCommand cmdArquivoCarregado = conn.CreateCommand();
            cmdArquivoCarregado.CommandText =
            " declare @tabela varchar(max) = 'D_Custo_Medio';" +
            " if (select count(arq_id) from S_ArquivoCarregado where Arq_Tabela = @tabela) = 0" +
            " insert into S_ArquivoCarregado" +
            " (Arq_ID, Arq_Nome, Arq_Tabela, Arq_Mensagem, Arq_DataCarga, Arq_Quantidade, Arq_Login)" +
            " values(1, '" + caminho + "', @tabela, 'Carga efetuada com sucesso.'," +
            " GETDATE(), ' " + linha.ToString() + "', REPLACE(SUSER_NAME(), 'ATRAME\\',''))" +
            " else" +
            " insert into S_ArquivoCarregado" +
            " (Arq_ID, Arq_Nome, Arq_Tabela, Arq_Mensagem, Arq_DataCarga, Arq_Quantidade, Arq_Login)" +
            " values(" + pegarID("D_Custo_Medio") + ", '" + caminho + "', @tabela, 'Carga efetuada com sucesso.'," +
            " GETDATE(), " + linha.ToString() + ", REPLACE(SUSER_NAME(), 'ATRAME\\',''))";

            conn.Open();
            SqlTransaction trA = null;
            trA = conn.BeginTransaction();
            cmdArquivoCarregado.Transaction = trA;
            cmdArquivoCarregado.ExecuteNonQuery();
            trA.Commit();

            conn.Close();
        }



        public void Produtos()
        {
            try
            {


                string filePath = null;
                //  string caminho = null;

                int linha = 1;
                int numRepetidos = 0, numCarregados = 0, numPendencias = 0;

                //            SqlConnection conn = new SqlConnection("Data Source=BRCAENRODRIGUES\\MSSQLSERVER01; Integrated Security=True; Initial Catalog=LAMPADA");

                //          filePath = @"C:\Base\info\produtos\produto.xlsx";

                filePath = caminho;

                FileInfo existingFile = new FileInfo(filePath);
                ExcelPackage package = new ExcelPackage(existingFile);
                ExcelWorksheet workSheet = package.Workbook.Worksheets[cmbPlanilha.SelectedIndex + 1];
                //ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
                StringBuilder conteudo = new StringBuilder();
                var lista = new List<String>();
                SqlCommand cmd = conn.CreateCommand();
                ArrayList Excel = new ArrayList();
                ArrayList SQL = new ArrayList();
                ArrayList repetido = new ArrayList();
                ArrayList carregado = new ArrayList();
                bool pendencia = false;

                // for (int o = workSheet.Dimension.Start.Row + 1; o <= workSheet.Dimension.End.Row; o++)
                //  {
                //    Excel.Add(workSheet.Cells[o, 1].Value.ToString());
                //    }

                if (conn.State.ToString() == "Closed")
                {
                    conn.Open();
                }


                SqlCommand cmdProc = conn.CreateCommand();
                SqlTransaction trProc = null;
                cmdProc.CommandText = "create or alter PROCEDURE [dbo].[SP_VERIFICA_PRODUTOS_REPETIDOS_CARREGADOR] @PROD VARCHAR(MAX) AS BEGIN IF NOT EXISTS(SELECT * FROM D_Produtos WHERE Pro_ID = @PROD)  BEGIN  RETURN 0; END ELSE  RETURN 1;  END ";
                trProc = conn.BeginTransaction();
                cmdProc.Transaction = trProc;
                cmdProc.ExecuteNonQuery();
                trProc.Commit();

                for (int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
                {
                    pendencia = false;

                    for (int j = workSheet.Dimension.Start.Column; j <= workSheet.Dimension.End.Column; j++)
                    {

                        //if ((j == 1) && (workSheet.Cells[i, 1].Value == null))
                        //{
                        //    MessageBox.Show("Test");
                        //    numPendencias++;
                        //    lblPendencia.Text = numPendencias.ToString();
                        //    lblPendencia.Refresh();
                        //}
                        //else
                        //{
                        if (j == 1 && (workSheet.Cells[i, j].Value == null || workSheet.Cells[linha, j].Value.ToString() == ""))
                        {


                            pendencia = true;
                            numPendencias++;
                            lblPendencia.Text = numPendencias.ToString();

                        }
                        else if ((j == 1))
                        {

                            SqlCommand cmdeProc = conn.CreateCommand();
                            cmdeProc.CommandType = CommandType.StoredProcedure;
                            cmdeProc.CommandText = "[SP_VERIFICA_PRODUTOS_REPETIDOS_CARREGADOR]";
                            cmdeProc.Parameters.Add("@PROD", SqlDbType.Int);
                            cmdeProc.Parameters["@PROD"].Direction = ParameterDirection.ReturnValue;
                            cmdeProc.Parameters.AddWithValue("@PROD", workSheet.Cells[i, j].Value);

                            if (conn.State.ToString() == "Closed")
                                conn.Open();

                            cmdeProc.ExecuteNonQuery();
                            int ret = Convert.ToInt32(cmdeProc.Parameters["@PROD"].Value);
                            conn.Close();

                            if (ret == 1)
                            {
                                numRepetidos++;
                                lblRepetido.Text = numRepetidos.ToString();
                                lblRepetido.Refresh();
                            }
                            else
                            {
                                numCarregados++;
                                lblCarregada.Text = numCarregados.ToString();
                                lblCarregada.Refresh();
                            }

                            conteudo.Append(" declare @pro_id varchar(max) = '" + workSheet.Cells[i, j].Value + "';");
                            conteudo.Append(Environment.NewLine);
                            conteudo.Append(" if  (select max(pro_id) from D_Produtos where pro_id = @pro_id) = (select (pro_id) from D_Produtos where pro_id = (@pro_id)) ");
                            conteudo.Append(Environment.NewLine);
                            conteudo.Append(" print 'OK' ");
                            conteudo.Append(Environment.NewLine);
                            conteudo.Append(" else ");
                            conteudo.Append(" INSERT INTO D_PRODUTOS " +
                                            " (PRO_ID, " +
                                            " PRO_DESCRICAO, " +
                                            " PRO_UND_ID, " +
                                            " PRO_NCM, " +
                                            " PRO_MARGEM, " +
                                            " Lin_Origem_ID, " +
                                            " [Arq_Origem_ID]) " +
                                            " VALUES ( ");
                            conteudo.Append("'" + workSheet.Cells[i, j].Value + "', ");
                        }
                        else if ((j == 3) && workSheet.Cells[i, j].Value != null)
                        {
                            Unidades unidade = (Unidades)System.Enum.Parse(typeof(Unidades), workSheet.Cells[i, j].Value.ToString());
                            conteudo.Append("'" + ((int)unidade).ToString() + "', ");
                        }
                        else if (j == workSheet.Dimension.End.Column)
                        {
                            conteudo.Append(workSheet.Cells[i, j].Value == null ? " '', '" + linha + "', " : " '" + workSheet.Cells[i, j].Value.ToString().Replace('\'', ' ') + "', '" + linha + "', ");
                            conteudo.Append(" " + pegarID("D_Produtos") + "  ");
                        }
                        else
                        {
                            conteudo.Append(workSheet.Cells[i, j].Value == null ? " NULL " : "'" + workSheet.Cells[i, j].Value.ToString().Replace('\'', ' ') + "', ");
                        }

                    }

                    if (i == workSheet.Dimension.End.Row)
                    {
                        conteudo.Append(" ) ");
                    }
                    else
                    {
                        conteudo.Append(")");
                        conteudo.Append(Environment.NewLine);
                    }

                    if (conn.State.ToString() == "Closed")
                    {
                        conn.Open();
                    }

                    if (pendencia == false)
                    {
                        // Clipboard.SetText(conteudo.ToString());
                        cmd.CommandText = conteudo.ToString();
                        //SqlTransaction trE = null;
                        //trE = conn.BeginTransaction();
                        //cmd.Transaction = trE;
                        //cmd.ExecuteNonQuery();
                        //trE.Commit();
                        fazTransacao(conn, cmd);
                    }

                    conteudo.Clear();
                    //        }
                }

                package.Dispose();


                if (numCarregados == 0)
                {
                    MessageBox.Show(new Form { TopMost = true }, "Nenhum registro de produtos carregado");
                }
                else
                {
                    MessageBox.Show(new Form { TopMost = true }, "Carregamento de " + numCarregados.ToString() + " registros de produtos realizados com sucesso");
                    SqlCommand cmdArquivoCarregado = conn.CreateCommand();
                    cmdArquivoCarregado.CommandText = gravaId(caminho, numCarregados, "D_Produtos");
                    fazTransacao(conn, cmdArquivoCarregado);
                }
            }
            catch (Exception)
            {

                throw;
            }
        }


        public void fazTransacao(SqlConnection conn, SqlCommand command)
        {

            if (conn.State.ToString() == "Closed")
                conn.Open();

            SqlTransaction trA = null;
            trA = conn.BeginTransaction();
            command.Transaction = trA;
            command.ExecuteNonQuery();
            trA.Commit();
            conn.Close();

        }

        public string gravaId(string caminho, int numCarregados, string tabela)
        {
            return " declare @tabela varchar(max) = '" + tabela + "';" +
                " if (select count(arq_id) from S_ArquivoCarregado where Arq_Tabela = @tabela) = 0" +
                " insert into S_ArquivoCarregado" +
                " (Arq_ID, Arq_Nome, Arq_Tabela, Arq_Mensagem, Arq_DataCarga, Arq_Quantidade, Arq_Login)" +
                " values(1, '" + caminho + "', @tabela, 'Carga efetuada com sucesso.'," +
                " GETDATE(), ' " + numCarregados.ToString() + "', REPLACE(SUSER_NAME(), 'ATRAME\\',''))" +
                " else" +
                " insert into S_ArquivoCarregado" +
                " (Arq_ID, Arq_Nome, Arq_Tabela, Arq_Mensagem, Arq_DataCarga, Arq_Quantidade, Arq_Login)" +
                " values( '" + pegarID(tabela) + "',  '" + caminho + "', @tabela, 'Carga efetuada com sucesso.'," +
                " GETDATE(), " + numCarregados.ToString() + ", REPLACE(SUSER_NAME(), 'ATRAME\\',''))" +
                "IF (OBJECT_ID('SP_VERIFICA_PRODUTOS_REPETIDOS_CARREGADOR') IS NOT NULL)  DROP PROCEDURE SP_VERIFICA_PRODUTOS_REPETIDOS_CARREGADOR  ";
        }

        private void button1_Click_2(object sender, EventArgs e)
        {

            if (cmbTabela.SelectedItem.Equals("D_Custo_Medio"))
            {
                Custo();
            }
            if (cmbTabela.SelectedItem.Equals("D_Clientes"))
            {
                Clientes();
            }
            if (cmbTabela.SelectedItem.Equals("D_Produtos"))
            {
                Produtos();
            }
            if (cmbTabela.SelectedItem.Equals("D_Inventario_Carga"))
            {
                Inventario();
            }
            if (cmbTabela.SelectedItem.Equals("D_Compras"))
            {
                Compras();
            }
            if (cmbTabela.SelectedItem.Equals("D_Vendas_Itens"))
            {
                Vendas();
            }
            if (cmbTabela.SelectedItem.Equals("D_Fornecedores"))
            {
                Fornecedores();
            }
            if (cmbTabela.SelectedItem.Equals("D_Insumo_Produto"))
            {
                Insumo_Produto();
            }
            if (cmbTabela.SelectedItem.Equals("D_Relacao_Carga"))
            {
                Relacao_Carga();
            }
            if (cmbTabela.SelectedItem.Equals("D_PIC"))
            {
                PIC();
            }

            //if (cmbTabela.SelectedItem == null || cmbTabela.SelectedIndex == 0 )
            //{
            //    MessageBox.Show("Selecione uma tabela para carregamento");
            //}

            //  lblCarregada.Text = "0";
            //  lblRepetido.Text = "0";
            //  lblPendencia.Text = "0";

        }

        enum Moedas
        {
            AFA = 5,
            ETB = 8,
            ARG = 10,
            THB = 15,
            PAB = 20,
            VEB = 25,
            BOB = 30,
            GHC = 35,
            CRC = 40,
            SVC = 45,
            NIC = 50,
            NIO = 51,
            DKK = 55,
            EEK = 57,
            SKK = 58,
            ISK = 60,
            NOK = 65,
            SEK = 70,
            CZK = 75,
            NCZ = 78,
            CZ = 79,
            CR = 80,
            RUR = 88,
            GMD = 90,
            DZD = 95,
            KWD = 100,
            BHD = 105,
            YD = 110,
            IQD = 115,
            DIN = 120,
            JOD = 125,
            LYD = 130,
            MKD = 132,
            SDD = 134,
            TND = 135,
            SDR = 138,
            MAD = 139,
            AED = 145,
            STD = 148,
            AUD = 150,
            BSD = 155,
            BMD = 160,
            CAD = 165,
            GYD = 170,
            BBD = 175,
            BZD = 180,
            BND = 185,
            KYD = 190,
            SGD = 195,
            FJD = 200,
            HKD = 205,
            TTD = 210,
            XCD = 215,
            ZWD = 217,
            USD = 220,
            JMD = 230,
            LRD = 235,
            M = 240,
            NZD = 245,
            SBD = 250,
            VND = 260,
            GRD = 270,
            CVE = 295,
            ESC = 315,
            TPE = 320,
            ANG = 325,
            AWG = 328,
            SRG = 330,
            NLG = 335,
            HUF = 345,
            BEF = 360,
            FBF = 361,
            BIF = 365,
            KMF = 368,
            XAF = 370,
            XPF = 380,
            DJF = 390,
            FRF = 395,
            GNF = 398,
            LUF = 400,
            MGF = 405,
            MF = 410,
            RWF = 420,
            CHF = 425,
            HTG = 440,
            PYG = 450,
            UAH = 460,
            JPY = 470,
            I = 480,
            GEL = 482,
            LVL = 485,
            ALL = 490,
            HNL = 495,
            SLL = 500,
            MDL = 503,
            ROL = 505,
            BGL = 510,
            CYP = 520,
            GIP = 530,
            EGP = 535,
            GBP = 540,
            FKP = 545,
            IEP = 550,
            IL = 555,
            LBP = 560,
            MTL = 565,
            SHP = 570,
            SYP = 575,
            LSD = 580,
            SZL = 585,
            ITL = 595,
            TRL = 600,
            LTL = 601,
            LSL = 603,
            AZM = 607,
            DEM = 610,
            BAM = 612,
            FMK = 615,
            MZM = 620,
            NGN = 630,
            AON = 635,
            YUM = 637,
            TWD = 640,
            MXN = 645,
            NCÇ = 651,
            PEN = 660,
            BTN = 665,
            MRO = 670,
            TOP = 680,
            MOP = 685,
            ADP = 690,
            ESP = 700,
            ARS = 706,
            B = 710,
            CLP = 715,
            COP = 720,
            CUP = 725,
            DOP = 730,
            PHP = 735,
            GWP = 738,
            MEX = 740,
            UYP = 745,
            BWP = 755,
            MWK = 760,
            ZMK = 765,
            GTQ = 770,
            MMK = 775,
            UAK = 776,
            PGK = 778,
            HRK = 779,
            LAK = 780,
            ZAR = 785,
            BRL = 790,
            CNY = 795,
            QAR = 800,
            OMR = 805,
            YER = 810,
            IRR = 815,
            SAR = 820,
            KHR = 825,
            MYR = 828,
            BYB = 829,
            RUB = 830,
            TJR = 835,
            MUR = 840,
            NPR = 845,
            SCR = 850,
            LKR = 855,
            INR = 860,
            IDR = 865,
            MVR = 870,
            PKR = 875,
            ILS = 880,
            S = 890,
            UZS = 893,
            ECS = 895,
            BDT = 905,
            WS = 910,
            WST = 911,
            KZT = 913,
            SIT = 914,
            MNT = 915,
            XEU = 918,
            VUV = 920,
            KPW = 925,
            KRW = 930,
            ATS = 940,
            TSH = 945,
            TZS = 946,
            KES = 950,
            UGX = 955,
            SOS = 960,
            ZRN = 970,
            PZN = 975,
            EUR = 978,
            CLRDA = 980,
            CLBULG = 982,
            CLGREC = 983,
            CLHUNG = 984,
            CLISR = 986,
            CLIUG = 988,
            CLPOL = 990,
            CLROM = 992,
            BUA = 995,
            FUA = 996,
            XAU = 998
        }
         
        enum Unidades
        {
            NULL = 11,
            X8 = 11,
            BD = 11,
            BL = 11,
            BR = 11,
            CAIXA = 11,
            CENTO = 11,
            CJ = 11,
            CM = 14,
            CT = 11,
            CX = 11,
            FD = 11,
            G = 22,
            GALAO = 11,
            H = 11,
            HORA = 11,
            HORAS = 11,
            JG = 10,
            KAR = 23,
            KG = 10,
            KI = 12,
            KILOMETROS = 14,
            L = 17,
            LATAO = 11,
            LB = 14,
            LITRO = 17,
            LITROS = 17,
            M = 14,
            M2 = 15,
            M3 = 16,
            METR2 = 15,
            METRO = 14,
            METROS3 = 16,
            ML = 12,
            MT = 14,
            ND = 11,
            OUTROS = 11,
            PA = 13,
            PAK = 23,
            PAL = 23,
            PARCELAS = 11,
            PC = 11,
            PECA = 11,
            PR = 13,
            PT = 11,
            RL = 11,
            SRV = 11,
            ST = 23,
            SV = 11,
            TO = 21,
            TON = 23,
            TR = 11,
            UN = 11,
            UNI = 23,
            UNIDA = 11,
            VOLUM = 11,
            x8 = 11,
            bd = 11,
            bl = 11,
            br = 11,
            caixa = 11,
            cento = 11,
            cj = 11,
            cm = 14,
            ct = 11,
            cx = 11,
            fd = 11,
            g = 22,
            galao = 11,
            h = 11,
            hora = 11,
            horas = 11,
            jg = 10,
            kar = 23,
            kg = 10,
            ki = 12,
            kilometros = 14,
            l = 17,
            latao = 11,
            lb = 14,
            litro = 17,
            litros = 17,
            m = 14,
            m2 = 15,
            m3 = 16,
            metr2 = 15,
            metro = 14,
            metros3 = 16,
            ml = 12,
            mt = 14,
            nd = 11,
            outros = 11,
            pa = 13,
            pak = 23,
            pal = 23,
            parcelas = 11,
            pc = 11,
            peca = 11,
            pr = 13,
            pt = 11,
            rl = 11,
            srv = 11,
            st = 23,
            sv = 11,
            to = 21,
            ton = 23,
            tr = 11,
            un = 11,
            uni = 23,
            unida = 11,
            volum = 11,
            Null = 11,
            Bd = 11,
            Bl = 11,
            Br = 11,
            Caixa = 11,
            Cento = 11,
            Cj = 11,
            Cm = 14,
            Ct = 11,
            Cx = 11,
            Fd = 11,
            Galao = 11,
            Hora = 11,
            Horas = 11,
            Jg = 10,
            Kar = 23,
            Kg = 10,
            Ki = 12,
            Kilometros = 14,
            Latao = 11,
            Lb = 14,
            Litro = 17,
            Litros = 17,
            Metr2 = 15,
            Metro = 14,
            Metros3 = 16,
            Ml = 12,
            Mt = 14,
            Nd = 11,
            Outros = 11,
            Pa = 13,
            Pak = 23,
            Pal = 23,
            Parcelas = 11,
            Pc = 11,
            Peca = 11,
            Pr = 13,
            Pt = 11,
            Rl = 11,
            Srv = 11,
            St = 23,
            Sv = 11,
            To = 21,
            Ton = 23,
            Tr = 11,
            Un = 11,
            Uni = 23,
            Unida = 11,
            Volum = 11
        }

        private void Form1_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void comboBoxServidor_SelectedIndexChanged(object sender, EventArgs e)
        {
            conexao = comboBoxServidor.Text;
            conn = new SqlConnection("Data Source=" + conexao + "; Integrated Security=True;");
            conn1 = new SqlConnection("Data Source=" + conexao + "; Integrated Security=True;");


            if (conn.State.ToString() == "Closed")
            {
                conn.Open();
            }
            System.Data.DataTable databases = conn.GetSchema("Databases");

            comboBoxBase.Items.Clear();

            foreach (DataRow database in databases.Rows)
            {
                databaseName = database.Field<String>("database_name");
                if (databaseName != "master" && databaseName != "tempdb" && databaseName != "model" && databaseName != "msdb")
                    comboBoxBase.Items.Add(databaseName);
            }
            conn.Close();
        }


        private void comboBoxServidor_Enter(object sender, EventArgs e)
        {
            comboBoxServidor.Items.Clear();
            string myServer = Environment.MachineName;

            System.Data.DataTable servers = SqlDataSourceEnumerator.Instance.GetDataSources();

            if (servers.Rows.Count > 0)
            {
                comboBoxServidor.Items.Add(servers.Rows[0]["ServerName"]);
                 
                for (int i = 0; i < servers.Rows.Count; i++)
                {

                    if (myServer == servers.Rows[i]["ServerName"].ToString())
                    {
                        if ((servers.Rows[i]["InstanceName"] as string) != null)
                        {
                            comboBoxServidor.Items.Add(servers.Rows[i]["ServerName"] + "\\" + servers.Rows[i]["InstanceName"]);
                        }
                        else
                        {
                            comboBoxServidor.Items.Add(servers.Rows[i]["ServerName"]);
                        }
                    }
                }
            }
            else
            {
                comboBoxServidor.Items.Add(myServer);
            }

        }

        private void comboBoxBase_SelectedIndexChanged(object sender, EventArgs e)
        {
            baseDeDados = comboBoxBase.Text;
            conn = new SqlConnection("Data Source=" + conexao + "; Integrated Security=True; Initial Catalog=" + baseDeDados);
            conn1 = new SqlConnection("Data Source=" + conexao + "; Integrated Security=True; Initial Catalog=" + baseDeDados);

        }

        private void button1_Click_3(object sender, EventArgs e)
        {

        }
        private void button1_Click_4(object sender, EventArgs e)
        { }

        private void button1_Click_5(object sender, EventArgs e)
        {


            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Compras");
                excel.Workbook.Worksheets.Add("Vendas");
                excel.Workbook.Worksheets.Add("Ordem");
                excel.Workbook.Worksheets.Add("Inventario");
                excel.Workbook.Worksheets.Add("Relacao");
                excel.Workbook.Worksheets.Add("Custo");
                excel.Workbook.Worksheets.Add("Clientes");
                excel.Workbook.Worksheets.Add("Produtos");
                excel.Workbook.Worksheets.Add("Fornecedores");
                excel.Workbook.Worksheets.Add("PIC");

                var comprasWorksheet = excel.Workbook.Worksheets["Compras"];
                var vendasWorksheet = excel.Workbook.Worksheets["Vendas"];
                var ordemWorksheet = excel.Workbook.Worksheets["Ordem"];
                var custoWorksheet = excel.Workbook.Worksheets["Custo"];
                var inventarioWorksheet = excel.Workbook.Worksheets["Inventario"];
                var relacaoWorksheet = excel.Workbook.Worksheets["Relacao"];
                var clientesWorksheet = excel.Workbook.Worksheets["Clientes"];
                var produtosWorksheet = excel.Workbook.Worksheets["Produtos"];
                var fornecedoresWorksheet = excel.Workbook.Worksheets["Fornecedores"];
                var picWorksheet = excel.Workbook.Worksheets["PIC"];

                List<string[]> headerRowCompras = new List<string[]>()
                    {
                        new string[] { "Código do Produto",   "Código Divisão",   "Código do Fornecedor", "Lançamento",   "Fatura",   "BL Data",  "Número da DI",    "Data da Importação",   "N da NF de Entrada",  "Serie",    "Data Entrada NF",  "CFOP NF Entrada",  "Data de Vencimento Média", "Dias", "Quantidade",   "Valor FOB (Moeda Estrangeira)",    "Código da Moeda Estrangeira",  "Frete",    "Seguro",   "Código Moeda Frete",   "Código Moeda Seguro",  "Imposto de Importação (Reais)",    "Icms", "Pis",  "Cofins", "Unidade", "CNPJ", "Incoterm",  "Id Fornecedor Seguro", "Id Fornecedor Frete" }
                    };

                List<string[]> headerRowVendas = new List<string[]>()
                    {
                      new string[] {"Código do Cliente",    "Número NF",    "Série NF", "Código da Divisão", "CFOP", "Data Emissão", "Data Vencimento",  "Prazo de Vencimento", "Item Nota Fiscal", "Código do Produto",    "Quantidade",   "Valor Venda sem o IPI (Reais)",  "Descontos Incondicionais", "ICMS", "PIS",  "COFINS",   "ISS",  "Comissão", "Frete",  "Seguro",   "Data de Embarque", "Código Moeda Estrangeira", "Valor em Moeda Estrangeira",   "Custo da Venda (CPV)" ,"RE", "CNPJ" }
                    };

                List<string[]> headerRowOrdem = new List<string[]>()
                    {
                      new string[] { "Código do Produto Acabado",   "Quantidade Produzida", "Unidade de Medida Produto Acabado",    "Código Matéria-Prima", "Quantidade Requisitada",   "Unidade de Medida Matéria-Prima",  "N da Ordem de Produção",  "Data Inínio",  "Data Fim", "CNPJ" }
                    };

                List<string[]> headerRowInventario = new List<string[]>()
                    {
                      new string[] { "Código do Produto",   "Data Inventário",  "Quantidade em Estoque", "Valor",   "Unidade de Medida", "CNPJ" }
                    };

                List<string[]> headerRowRelacao = new List<string[]>()
                    {
                      new string[] { "Produto Acabado", "Matéria Prima", "Quantidade Produzida", "Quantidade Requisitada", "Relacao", "Tipo Relação" }
                    };

                List<string[]> headerRowClientes = new List<string[]>()
                    {
                      new string[] { "Código do Cliente", "Nome", "Código do País","Vínculo",  "Data Inicio",  "Data Fim", "CNPJ"}
                    };

                List<string[]> headerRowProdutos = new List<string[]>()
                    {
                      new string[] {"Código do Produto",    "Descrição",    "Unidade de Medida",    "Classificação Fiscal (NCM)", "Margem"}
                    };

                List<string[]> headerRowFornecedores = new List<string[]>()
                    {
                      new string[] { "Código do Fornecedor", "Nome", "Código do País","Vínculo", "Data Inicio",  "Data Fim", "CNPJ"}
                    };

                List<string[]> headerRowCusto = new List<string[]>()
                    {
                      new string[] {"Código do Produto", "Mês", "Ano", "Custo Médio Unitário", "CNPJ" }
                    };

                List<string[]> headerRowPIC = new List<string[]>()
                    {
                      new string[] {"Fornecedor Código",   "Fornecedor Pais", "Pro Código",   "Qtde", "Vl Moeda Bruto",   "Cod Moeda",    "Data Emissao",   "Data Venc",  "Cli Código",   "Cli Pais", "Doc Oper", "Dias", "CFOP", "Imposto Intern",   "FOB Moeda",    "IPI",  "ICMS", "Pis",  "Cofins",   "Frete",    "Seguro"}
                    };

                string headerRangeFornecedores = "A1:" + Char.ConvertFromUtf32(headerRowFornecedores[0].Length + 64) + "1";
                fornecedoresWorksheet.Cells[headerRangeFornecedores].LoadFromArrays(headerRowFornecedores);
                fornecedoresWorksheet.Cells[headerRangeFornecedores].Style.Font.Bold = true;
                fornecedoresWorksheet.Column(1).AutoFit();

                string headerRangeCompras = "A1:AB1";
                comprasWorksheet.Cells[headerRangeCompras].LoadFromArrays(headerRowCompras);
                comprasWorksheet.Cells[headerRangeCompras].Style.Font.Bold = true;
                comprasWorksheet.Column(1).AutoFit();

                string headerRangeInventario = "A1:" + Char.ConvertFromUtf32(headerRowInventario[0].Length + 64) + "1";
                inventarioWorksheet.Cells[headerRangeInventario].LoadFromArrays(headerRowInventario);
                inventarioWorksheet.Cells[headerRangeInventario].Style.Font.Bold = true;
                inventarioWorksheet.Column(1).AutoFit();

                string headerRangeVendas = "A1:" + Char.ConvertFromUtf32(headerRowVendas[0].Length + 64) + "1";
                vendasWorksheet.Cells[headerRangeVendas].LoadFromArrays(headerRowVendas);
                vendasWorksheet.Cells[headerRangeVendas].Style.Font.Bold = true;
                vendasWorksheet.Column(1).AutoFit();

                string headerRangeOrdem = "A1:" + Char.ConvertFromUtf32(headerRowOrdem[0].Length + 64) + "1";
                ordemWorksheet.Cells[headerRangeOrdem].LoadFromArrays(headerRowOrdem);
                ordemWorksheet.Cells[headerRangeOrdem].Style.Font.Bold = true;
                ordemWorksheet.Column(1).AutoFit();

                string headerRangeClientes = "A1:" + Char.ConvertFromUtf32(headerRowClientes[0].Length + 64) + "1";
                clientesWorksheet.Cells[headerRangeClientes].LoadFromArrays(headerRowClientes);
                clientesWorksheet.Cells[headerRangeClientes].Style.Font.Bold = true;
                clientesWorksheet.Column(1).AutoFit();

                string headerRangeProdutos = "A1:" + Char.ConvertFromUtf32(headerRowProdutos[0].Length + 64) + "1";
                produtosWorksheet.Cells[headerRangeProdutos].LoadFromArrays(headerRowProdutos);
                produtosWorksheet.Cells[headerRangeProdutos].Style.Font.Bold = true;
                produtosWorksheet.Column(1).AutoFit();

                string headerRangeRelacao = "A1:" + Char.ConvertFromUtf32(headerRowRelacao[0].Length + 64) + "1";
                relacaoWorksheet.Cells[headerRangeRelacao].LoadFromArrays(headerRowRelacao);
                relacaoWorksheet.Cells[headerRangeRelacao].Style.Font.Bold = true;
                relacaoWorksheet.Column(1).AutoFit();

                string headerRangeCusto = "A1:" + Char.ConvertFromUtf32(headerRowCusto[0].Length + 64) + "1";
                custoWorksheet.Cells[headerRangeCusto].LoadFromArrays(headerRowCusto);
                custoWorksheet.Cells[headerRangeCusto].Style.Font.Bold = true;
                custoWorksheet.Column(1).AutoFit();

                string headerRangePIC = "A1:" + Char.ConvertFromUtf32(headerRowPIC[0].Length + 64) + "1";
                picWorksheet.Cells[headerRangePIC].LoadFromArrays(headerRowPIC);
                picWorksheet.Cells[headerRangePIC].Style.Font.Bold = true;
                picWorksheet.Column(1).AutoFit();

                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "Excel|*.xlsx";
                saveFileDialog1.Title = "Salvar Excel";
                saveFileDialog1.FileName = "Modelo Carregamento Dados TPS.xlsx";
                saveFileDialog1.ShowDialog();

                // If the file name is not an empty string open it for saving.  
                if (saveFileDialog1.FileName != "")
                {
                    // Saves the Image via a FileStream created by the OpenFile method.  
                    System.IO.FileStream fs = (System.IO.FileStream)saveFileDialog1.OpenFile();
                    excel.SaveAs(fs);
                }
                excel.Dispose();
            }

        }

        private void cmbPlanilha_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filePath = caminho;
            FileInfo existingFile = new FileInfo(filePath);
            ExcelPackage package = new ExcelPackage(existingFile);
            ExcelWorksheet workSheet = package.Workbook.Worksheets[cmbPlanilha.SelectedIndex + 1];
            StringBuilder conteudo = new StringBuilder();
            var lista = new List<String>();
            SqlCommand cmd = conn.CreateCommand();
            int linha = 1;

            lblTotal.Text = (workSheet.Dimension.End.Row - 1).ToString();
            lblTotal.Refresh();
        }

        private void cmbTabela_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblRepetido.Text = "0";
            lblPendencia.Text = "0";
            lblCarregada.Text = "0";
        }
    }
}
