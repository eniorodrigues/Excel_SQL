using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace testeCampos
{
    class Compras
    {

//        using System;
//using System.Collections.Generic;
//using System.Data;
//using System.Text;
//using System.Windows.Forms;
//using System.Data.OleDb;
//using System.Data.SqlClient;
//using System.Data.Common;
//using Excel = Microsoft.Office.Interop.Excel;
//using System.IO;
//using System.Data.Sql;
//using System.Runtime.InteropServices;
//using System.Configuration;
//using Microsoft.Office.Interop.Excel;
//using ExcelIt = Microsoft.Office.Interop.Excel;
//using System.Reflection;
//using System.Drawing;
//using testeCampos;
//using System.ComponentModel;
//using System.Text.RegularExpressions;
//using System.Threading;
//using OfficeOpenXml;
//using System.Linq;
//using System.ComponentModel.DataAnnotations;
//using System.Globalization;

//namespace testeExcel
//    {
//        public partial class Form1 : Form
//        {
//            public Form1()
//            {
//                InitializeComponent();
//            }

//            public static string path;
//            public static string excelConnectionString;
//            public string[] files;
//            public string conexao;
//            public string baseDeDados;
//            public string tabela;
//            public string caminho;
//            public string directoryPath;
//            private static Excel.Application MyApp = null;
//            public List<string> filesAdionado = new List<string>();
//            public List<string> colunas = new List<string>();
//            public List<string> colunasCreate = new List<string>();
//            public string tipoArquivo;
//            Stream myStream = null;
//            string nomeSheet;
//            StringBuilder camposDataGrid = new StringBuilder();
//            public List<String> itemsDataGrid = new List<String>();
//            public ClientesTeste clientesTeste = new ClientesTeste();
//            public FornecedoresTeste fornecedores = new FornecedoresTeste();
//            public ProdutoTeste produtosTeste = new ProdutoTeste();
//            public Inventario inventario = new Inventario();
//            public InsumoProduto insumoProduto = new InsumoProduto();
//            public Vendas vendas = new Vendas();
//            public Compras compras = new Compras();
//            public Relacao relacao = new Relacao();
//            public SqlConnection conn = null;
//            public bool checado;

//            private void button1_Click(object sender, EventArgs e)
//            {
//                btnTarefaIndeterminada.Enabled = false;
//                btnTarefaIndeterminada.Enabled = false;

//                progressBar1.Style = ProgressBarStyle.Blocks;
//                progressBar1.Value = 0;

//                backgroundWorker1.RunWorkerAsync();
//            }

//            private void buttonAbrir_Click(object sender, EventArgs e)
//            {
//                SqlConnection conn = new SqlConnection(@"Data Source=BRCAENRODRIGUES\SQLEXPRESS01; Initial Catalog=LAMPADA; Integrated Security=True");
//                OpenFileDialog openFileDialog1 = new OpenFileDialog();
//                openFileDialog1.InitialDirectory = "C:\\";
//                openFileDialog1.Filter = "Csv files (*.csv*)|*.csv*|Excel files (*.xls*)|*.xls*";
//                openFileDialog1.FilterIndex = 2;
//                openFileDialog1.RestoreDirectory = true;
//                openFileDialog1.Multiselect = true;

//                if (openFileDialog1.ShowDialog() == DialogResult.OK)
//                {
//                    try
//                    {
//                        if ((myStream = openFileDialog1.OpenFile()) != null)
//                        {
//                            using (myStream)
//                            {
//                                caminho = openFileDialog1.FileName;
//                                directoryPath = Path.GetDirectoryName(openFileDialog1.FileName);
//                                files = (openFileDialog1.SafeFileNames);

//                                foreach (string file in files)
//                                {
//                                    filesAdionado.Add(file);
//                                    listBox1.Items.Add(file);
//                                }
//                                carregaLinhas();
//                            }
//                        }
//                    }
//                    catch (Exception ex)
//                    {
//                        MessageBox.Show(ex.Message);
//                    }
//                }
//            }

//            public void carregaLinhas()
//            {
//                label2.Text = caminho;

//                MyApp = new Excel.Application();
//                excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + caminho + ";Extended Properties=Excel 12.0;";
//                MyApp.Workbooks.Add("");
//                MyApp.Workbooks.Add(caminho);
//                SqlTransaction trAx = null;

//                for (int i = 1; i <= MyApp.Workbooks[2].Worksheets.Count; i++)
//                {
//                    comboBox2.Items.Add(MyApp.Workbooks[2].Worksheets[i].Name);
//                }
//            }

//            private void Form1_Load(object sender, EventArgs e)
//            {
//                conn = new SqlConnection(@"Data Source=BRCAENRODRIGUES\SQLEXPRESS01; Initial Catalog=LAMPADA; Integrated Security=True");
//            }

//            static System.Data.DataTable ConvertListToDataTable(List<string> list)
//            {
//                System.Data.DataTable table = new System.Data.DataTable();
//                for (int i = 0; i < 1; i++)
//                {
//                    table.Columns.Add();
//                    table.Columns[0].ColumnName = "Campos Excel";
//                }
//                foreach (var array in list)
//                {
//                    table.Rows.Add(array);
//                }
//                return table;
//            }

//            private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
//            {
//                System.Text.StringBuilder builder = new System.Text.StringBuilder();
//                builder.Append(" INSERT INTO[dbo].[D_Compras] " +
//                    "([Cmp_Pro_ID] " +
//                    ",[Cmp_Cod_Divisao]" +
//                    ",[Cmp_For_ID]" +
//                    ",[Cmp_Lanc_Cont]" +
//                    ",[Cmp_Fat_Coml]" +
//                    ",[Cmp_BL_DT]" +
//                    ",[Cmp_DI_ID]" +
//                    ",[Cmp_DI_DT_Emissao]" +
//                    ",[Cmp_NF_Entrada]" +
//                    ",[Cmp_NF_Serie]" +
//                    ",[Cmp_NF_DT]" +
//                    ",[Cmp_CFOP]" +
//                    ",[Cmp_DI_DT_Vencimento]" +
//                    ",[Cmp_DI_Dias]" +
//                    ",[Cmp_Qtde]" +
//                    ",[Cmp_Valor_Fob]" +
//                    ",[Cmp_Cod_Moeda]" +
//                    ",[Cmp_Vl_Frete_Moeda]" +
//                    ",[Cmp_VL_Seguro_Moeda]" +
//                    ",[Cmp_Cod_Moeda_Frete]" +
//                    ",[Cmp_Cod_Moeda_Seguro]" +
//                    ",[Cmp_Imposto_Import]" +
//                    ",[Cmp_Incoterm]) " +
//                     Environment.NewLine + " VALUES ( ");
//                for (int r = 2; r <= MyApp.Workbooks[2].Worksheets[1].UsedRange.Rows.Count; r++)
//                {
//                    for (int k = 1; k <= MyApp.Workbooks[2].Worksheets[comboBox2.SelectedIndex + 1].UsedRange.Columns.Count; k++)
//                    {
//                        if (k == MyApp.Workbooks[2].Worksheets[comboBox2.SelectedIndex + 1].UsedRange.Columns.Count)
//                        {
//                            builder.Append(MyApp.Workbooks[2].Worksheets[1].Cells[r, k].Value2 == null ? " NULL " : "'" + MyApp.Workbooks[2].Worksheets[1].Cells[r, k].Value2.ToString().Replace(',', '.') + "' ");
//                        }
//                        else
//                        {
//                            builder.Append(MyApp.Workbooks[2].Worksheets[1].Cells[r, k].Value2 == null ? " NULL, " : "'" + MyApp.Workbooks[2].Worksheets[1].Cells[r, k].Value2.ToString().Replace(',', '.') + "', ");
//                        }
//                    }
//                    if (r == MyApp.Workbooks[2].Worksheets[1].UsedRange.Rows.Count)
//                    {
//                        builder.Append(") ");
//                    }
//                    else
//                    {
//                        builder.Append("),");
//                        builder.Append(Environment.NewLine);
//                        builder.Append(" (");
//                    }
//                }

//                Clipboard.SetText(builder.ToString());

//                MessageBox.Show("Inclusão concluída");

//                System.Data.DataTable table = ConvertListToDataTable(colunas);
//                dataGridView2.DataSource = table;
//                label5.Text = comboBox2.SelectedItem.ToString();
//            }

//            private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
//            {

//            }


//            //private void TarefaLonga(int p)
//            //{
//            //    for (int i = 0; i <= 10; i++)
//            //    {
//            //        // faz a thread dormir por "p" milissegundos a cada passagem do loop
//            //        Thread.Sleep(p);
//            //        label2.BeginInvoke(
//            //           new System.Action(() =>
//            //           {
//            //               label2.Text = "Tarefa: " + i.ToString() + " concluída";
//            //           }
//            //        ));
//            //    }
//            //}

//            private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
//            {
//                //for (int i = 0; i < 100; i++)//representa uma tarefa com 100 processos.
//                //{
//                //Executa o método longo 100 vezes.
//                TarefaLonga(1);
//                //incrementa o progresso do backgroundWorker 
//                //a cada passagem do loop.
//                //this.backgroundWorker1.ReportProgress(i);
//                this.backgroundWorker1.ReportProgress(1);
//                //Verifica se houve uma requisição para cancelar a operação.
//                if (backgroundWorker1.CancellationPending)
//                {
//                    //se sim, define a propriedade Cancel para true
//                    //para que o evento WorkerCompleted saiba que a tarefa foi cancelada.
//                    e.Cancel = true;

//                    //zera o percentual de progresso do backgroundWorker1.
//                    backgroundWorker1.ReportProgress(0);
//                    return;
//                }
//                //}
//                //Finalmente, caso tudo esteja ok, finaliza
//                //o progresso em 100%.
//                backgroundWorker1.ReportProgress(100);
//            }

//            /// <summary>
//            /// Aqui implementamos o que desejamos fazer enquanto o progresso
//            /// da tarefa é modificado,[incrementado].
//            /// </summary>
//            private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
//            {
//                //Incrementa o valor da progressbar com o valor
//                //atual do progresso da tarefa.
//                progressBar1.Value = e.ProgressPercentage;

//                //informa o percentual na forma de texto.
//                label1.Text = e.ProgressPercentage.ToString() + "%";
//            }

//            /// <summary>
//            /// Após a tarefa ser concluida, esse metodo e chamado para
//            /// implementar o que deve ser feito imediatamente após a conclusão da tarefa.
//            /// </summary>
//            private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
//            {
//                if (e.Cancelled)
//                {
//                    //caso a operação seja cancelada, informa ao usuario.
//                    label2.Text = "Operação Cancelada pelo Usuário!";

//                    //habilita o Botao cancelar
//                    btnCancelar.Enabled = true;
//                    //limpa a label
//                    label1.Text = string.Empty;
//                }
//                else if (e.Error != null)
//                {
//                    //informa ao usuario do acontecimento de algum erro.
//                    label2.Text = "Aconteceu um erro durante a execução do processo!";
//                }
//                else
//                {
//                    //informa que a tarefa foi concluida com sucesso.
//                    label2.Text = "Tarefa Concluida com sucesso!";
//                }
//                //habilita os botões.
//                btnTarefaDeterminada.Enabled = true;
//                btnTarefaIndeterminada.Enabled = true;
//            }

//            private void btnCancelar_Click(object sender, EventArgs e)
//            {
//                //Cancelamento da tarefa com fim determinado [backgroundWorker1]
//                if (backgroundWorker1.IsBusy)//se o backgroundWorker1 estiver ocupado
//                {
//                    // notifica a thread que o cancelamento foi solicitado.
//                    // Cancela a tarefa DoWork 
//                    backgroundWorker1.CancelAsync();
//                }

//                //Cancelamento da tarefa com fim indeterminado [bgWorkerIndeterminada]
//                if (bgWorkerIndeterminada.IsBusy)
//                {
//                    // notifica a thread que o cancelamento foi solicitado.
//                    // Cancela a tarefa DoWork 
//                    bgWorkerIndeterminada.CancelAsync();
//                }

//                //desabilita o botão cancelar.
//                btnCancelar.Enabled = false;
//                label1.Text = "Cancelando...";
//            }

//            private void bgWorkerIndeterminada_DoWork(object sender, DoWorkEventArgs e)
//            {
//                //executa a tarefa a primeira vez
//                TarefaLonga(1);
//                //Verifica se houve uma requisição para cancelar a operação.
//                if (bgWorkerIndeterminada.CancellationPending)
//                {
//                    //se sim, define a propriedade Cancel para true
//                    //para que o evento WorkerCompleted saiba que a tarefa foi cancelada.
//                    e.Cancel = true;
//                    return;
//                }

//                //executa a tarefa pela segunda vez
//                //TarefaLonga(200);
//                //if (bgWorkerIndeterminada.CancellationPending)
//                //{
//                //    //se sim, define a propriedade Cancel para true
//                //    //para que o evento WorkerCompleted saiba que a tarefa foi cancelada.
//                //    e.Cancel = true;
//                //    return;
//                //}
//            }

//            private void bgWorkerIndeterminada_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
//            {
//                //Caso cancelado...
//                if (e.Cancelled)
//                {
//                    // reconfigura a progressbar para o padrao.
//                    progressBar1.MarqueeAnimationSpeed = 0;
//                    progressBar1.Style = ProgressBarStyle.Blocks;
//                    progressBar1.Value = 0;

//                    //caso a operação seja cancelada, informa ao usuario.
//                    label2.Text = "Operação Cancelada pelo Usuário!";

//                    //habilita o botao cancelar
//                    btnCancelar.Enabled = true;
//                    //limpa a label
//                    label1.Text = string.Empty;
//                }
//                else if (e.Error != null)
//                {
//                    //informa ao usuario do acontecimento de algum erro.
//                    label2.Text = "Aconteceu um erro durante a execução do processo!";

//                    // reconfigura a progressbar para o padrao.
//                    progressBar1.MarqueeAnimationSpeed = 0;
//                    progressBar1.Style = ProgressBarStyle.Blocks;
//                    progressBar1.Value = 0;
//                }
//                else
//                {
//                    //informa que a tarefa foi concluida com sucesso.
//                    label2.Text = "Tarefa Concluida com sucesso!";

//                    //Carrega todo progressbar.
//                    progressBar1.MarqueeAnimationSpeed = 0;
//                    progressBar1.Style = ProgressBarStyle.Blocks;
//                    progressBar1.Value = 100;
//                    label1.Text = progressBar1.Value.ToString() + "%";
//                }
//                //habilita os botões.
//                btnTarefaDeterminada.Enabled = true;
//                btnTarefaIndeterminada.Enabled = true;
//            }

//            private void btnTarefaIndeterminada_Click(object sender, EventArgs e)
//            {
//                //desabilita os botões enquanto a tarefa é executada.
//                btnTarefaDeterminada.Enabled = false;
//                btnTarefaIndeterminada.Enabled = false;
//                bgWorkerIndeterminada.RunWorkerAsync();

//                //define a progressBar para Marquee
//                progressBar1.Style = ProgressBarStyle.Marquee;
//                progressBar1.MarqueeAnimationSpeed = 5;

//                //informa que a tarefa esta sendo executada.
//                label1.Text = "Processando...";
//            }

//            private void button1_Click_1(object sender, EventArgs e)
//            {

//            }

//            private void bgWorkerIndeterminada_ProgressChanged(object sender, ProgressChangedEventArgs e)
//            {

//            }

//            public void TarefaLonga(int coisa)
//            {
//                string filePath = "C:\\Base\\Book1.xlsx";

//                // Abrindo, modificando meu arquivo e salvando
//                ExcelPackage package = new ExcelPackage(new FileInfo(filePath));
//                ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
//                DateTime hoje = DateTime.Now;

//                int lastRow = workSheet.Dimension.End.Row;

//                //workSheet.Cells[1, 1].Value = hoje.ToString(); // sobrescrevendo primeira linha e coluna
//                //workSheet.Cells[lastRow + 1, 1].Value = hoje.ToString(); // inserindo data de hoje apos ultima linha, 1a coluna

//                package.Save();
//                package.Dispose();

//                // Abrindo meu arquivo para varrer o conteudo da primeira planilha e exibir na minha página
//                package = new ExcelPackage(new FileInfo(filePath));
//                workSheet = package.Workbook.Worksheets.First();
//                StringBuilder conteudo = new StringBuilder();

//                conteudo.Append(" INSERT INTO[dbo].[D_Compras] " +
//                   "([Cmp_Pro_ID] " +
//                   ",[Cmp_Cod_Divisao]" +
//                   ",[Cmp_For_ID]" +
//                   ",[Cmp_Lanc_Cont]" +
//                   ",[Cmp_Fat_Coml]" +
//                   ",[Cmp_BL_DT]" +
//                   ",[Cmp_DI_ID]" +
//                   ",[Cmp_DI_DT_Emissao]" +
//                   ",[Cmp_NF_Entrada]" +
//                   ",[Cmp_NF_Serie]" +
//                   ",[Cmp_NF_DT]" +
//                   ",[Cmp_CFOP]" +
//                   ",[Cmp_DI_DT_Vencimento]" +
//                   ",[Cmp_DI_Dias]" +
//                   ",[Cmp_Qtde]" +
//                   ",[Cmp_Valor_Fob]" +
//                   ",[Cmp_Cod_Moeda]" +
//                   ",[Cmp_Vl_Frete_Moeda]" +
//                   ",[Cmp_VL_Seguro_Moeda]" +
//                   ",[Cmp_Cod_Moeda_Frete]" +
//                   ",[Cmp_Cod_Moeda_Seguro]" +
//                   ",[Cmp_Imposto_Import]" +
//                   ",[Cmp_Incoterm]) " +
//                    Environment.NewLine + " VALUES ( ");

//                for (int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
//                {
//                    for (int j = workSheet.Dimension.Start.Column; j <= workSheet.Dimension.End.Column; j++)
//                    {
//                        ///ultima coluna
//                        if (j == workSheet.Dimension.End.Column)
//                        {
//                            //MessageBox.Show("ultima coluna");
//                            conteudo.Append(workSheet.Cells[i, j].Value == null ? " NULL " : "'" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "' ");
//                        }
//                        else
//                        {
//                            //demais colunas a partir da segunda
//                            //data en-us
//                            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US")
//                            {
//                                DateTimeFormat = { YearMonthPattern = "yyyy-mm-dd" }
//                            };
//                            //caso número seja nula colocar zero
//                            if ((j == 19 || j == 18 || j == 15 || j == 16 || j == 22) && workSheet.Cells[i, j].Value == null)
//                            {
//                                conteudo.Append(" " + 0 + ", ");
//                            }
//                            //caso número tirar  aspas simples
//                            else if ((j == 19 || j == 18 || j == 15 || j == 16 || j == 22) && workSheet.Cells[i, j].Value != null)
//                            {
//                                conteudo.Append(workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + ", ");
//                            }
//                            //convert cadata caso convertido em integer
//                            else if ((j == 6 || j == 8 || j == 11 || j == 13) && workSheet.Cells[i, j].Value != null && workSheet.Cells[i, j].Value.GetType().Name.ToString() != "DateTime")
//                            {
//                                DateTime dt = DateTime.FromOADate(Convert.ToInt64(workSheet.Cells[i, j].Value));
//                                conteudo.Append("'" + dt + "', ");
//                            }
//                            //faz o depara do código de moeda
//                            else if ((j == 21 || j == 20 || j == 17) && workSheet.Cells[i, j].Value != null)
//                            {
//                                Moedas moeda = (Moedas)System.Enum.Parse(typeof(Moedas), workSheet.Cells[i, j].Value.ToString());
//                                conteudo.Append("'" + ((int)moeda).ToString() + "', ");
//                            }
//                            else
//                            {
//                                if ((workSheet.Cells[i, j].Value == null ? " NULL " : workSheet.Cells[i, j].Value.GetType().Name.ToString()) == "DateTime")
//                                {
//                                    workSheet.Cells[i, j].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.YearMonthPattern;
//                                    conteudo.Append(workSheet.Cells[i, j].Value == null ? " NULL, " : "'" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "', ");
//                                }
//                                else
//                                {
//                                    conteudo.Append(workSheet.Cells[i, j].Value == null ? " NULL, " : "'" + workSheet.Cells[i, j].Value.ToString().Replace(',', '.') + "', ");
//                                }
//                            }
//                        }
//                    }

//                    if (i == workSheet.Dimension.End.Row)
//                    {
//                        conteudo.Append(") ");
//                    }
//                    else
//                    {
//                        conteudo.Append("),");
//                        conteudo.Append(Environment.NewLine);
//                        conteudo.Append(" (");
//                    }
//                }

//                Clipboard.SetText(conteudo.ToString());

//                package.Dispose();
//            }

//            private void button1_Click_2(object sender, EventArgs e)
//            {
//                TarefaLonga(1);
//            }


//            enum Moedas
//            {
//                AFA = 5,
//                ETB = 8,
//                ARG = 10,
//                THB = 15,
//                PAB = 20,
//                VEB = 25,
//                BOB = 30,
//                GHC = 35,
//                CRC = 40,
//                SVC = 45,
//                NIC = 50,
//                NIO = 51,
//                DKK = 55,
//                EEK = 57,
//                SKK = 58,
//                ISK = 60,
//                NOK = 65,
//                SEK = 70,
//                CZK = 75,
//                NCZ = 78,
//                CZ = 79,
//                CR = 80,
//                RUR = 88,
//                GMD = 90,
//                DZD = 95,
//                KWD = 100,
//                BHD = 105,
//                YD = 110,
//                IQD = 115,
//                DIN = 120,
//                JOD = 125,
//                LYD = 130,
//                MKD = 132,
//                SDD = 134,
//                TND = 135,
//                SDR = 138,
//                MAD = 139,
//                AED = 145,
//                STD = 148,
//                AUD = 150,
//                BSD = 155,
//                BMD = 160,
//                CAD = 165,
//                GYD = 170,
//                BBD = 175,
//                BZD = 180,
//                BND = 185,
//                KYD = 190,
//                SGD = 195,
//                FJD = 200,
//                HKD = 205,
//                TTD = 210,
//                XCD = 215,
//                ZWD = 217,
//                USD = 220,
//                JMD = 230,
//                LRD = 235,
//                M = 240,
//                NZD = 245,
//                SBD = 250,
//                VND = 260,
//                GRD = 270,
//                CVE = 295,
//                ESC = 315,
//                TPE = 320,
//                ANG = 325,
//                AWG = 328,
//                SRG = 330,
//                NLG = 335,
//                HUF = 345,
//                BEF = 360,
//                FBF = 361,
//                BIF = 365,
//                KMF = 368,
//                XAF = 370,
//                XPF = 380,
//                DJF = 390,
//                FRF = 395,
//                GNF = 398,
//                LUF = 400,
//                MGF = 405,
//                MF = 410,
//                RWF = 420,
//                CHF = 425,
//                HTG = 440,
//                PYG = 450,
//                UAH = 460,
//                JPY = 470,
//                I = 480,
//                GEL = 482,
//                LVL = 485,
//                ALL = 490,
//                HNL = 495,
//                SLL = 500,
//                MDL = 503,
//                ROL = 505,
//                BGL = 510,
//                CYP = 520,
//                GIP = 530,
//                EGP = 535,
//                GBP = 540,
//                FKP = 545,
//                IEP = 550,
//                IL = 555,
//                LBP = 560,
//                MTL = 565,
//                SHP = 570,
//                SYP = 575,
//                LSD = 580,
//                SZL = 585,
//                ITL = 595,
//                TRL = 600,
//                LTL = 601,
//                LSL = 603,
//                AZM = 607,
//                DEM = 610,
//                BAM = 612,
//                FMK = 615,
//                MZM = 620,
//                NGN = 630,
//                AON = 635,
//                YUM = 637,
//                TWD = 640,
//                MXN = 645,
//                NCÇ = 651,
//                PEN = 660,
//                BTN = 665,
//                MRO = 670,
//                TOP = 680,
//                MOP = 685,
//                ADP = 690,
//                ESP = 700,
//                ARS = 706,
//                B = 710,
//                CLP = 715,
//                COP = 720,
//                CUP = 725,
//                DOP = 730,
//                PHP = 735,
//                GWP = 738,
//                MEX = 740,
//                UYP = 745,
//                BWP = 755,
//                MWK = 760,
//                ZMK = 765,
//                GTQ = 770,
//                MMK = 775,
//                UAK = 776,
//                PGK = 778,
//                HRK = 779,
//                LAK = 780,
//                ZAR = 785,
//                BRL = 790,
//                CNY = 795,
//                QAR = 800,
//                OMR = 805,
//                YER = 810,
//                IRR = 815,
//                SAR = 820,
//                KHR = 825,
//                MYR = 828,
//                BYB = 829,
//                RUB = 830,
//                TJR = 835,
//                MUR = 840,
//                NPR = 845,
//                SCR = 850,
//                LKR = 855,
//                INR = 860,
//                IDR = 865,
//                MVR = 870,
//                PKR = 875,
//                ILS = 880,
//                S = 890,
//                UZS = 893,
//                ECS = 895,
//                BDT = 905,
//                WS = 910,
//                WST = 911,
//                KZT = 913,
//                SIT = 914,
//                MNT = 915,
//                XEU = 918,
//                VUV = 920,
//                KPW = 925,
//                KRW = 930,
//                ATS = 940,
//                TSH = 945,
//                TZS = 946,
//                KES = 950,
//                UGX = 955,
//                SOS = 960,
//                ZRN = 970,
//                PZN = 975,
//                EUR = 978,
//                CLRDA = 980,
//                CLBULG = 982,
//                CLGREC = 983,
//                CLHUNG = 984,
//                CLISR = 986,
//                CLIUG = 988,
//                CLPOL = 990,
//                CLROM = 992,
//                BUA = 995,
//                FUA = 996,
//                XAU = 998
//            }

//        }
//    }

}
}
