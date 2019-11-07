using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace testeCampos
{
    class Temp
    {

//        using Microsoft.Office.Interop.Excel;
//using OfficeOpenXml;
//using System;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.Data;
//using System.Data.SqlClient;
//using System.Drawing;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;
//using Microsoft.Office;
//using System.Collections;
//using System.Threading;
//using System.Globalization;

//namespace BOSTA
//    {
//        public partial class Form1 : Form
//        {

//            SqlConnection conn = new SqlConnection("Data Source=BRCAENRODRIGUES\\MSSQLSERVER01; Integrated Security=True;MultipleActiveResultSets=true; Initial Catalog=LAMPADA");

//            public Form1()
//            {
//                InitializeComponent();
//            }

//            private void button1_Click(object sender, EventArgs e)
//            {

//                SqlCommand cmd = conn.CreateCommand();
//                cmd.CommandType = CommandType.StoredProcedure;
//                cmd.CommandText = "[SP_JACA_PRODUTOS]";
//                cmd.Parameters.Add("@PROD", SqlDbType.Int);
//                cmd.Parameters["@PROD"].Direction = ParameterDirection.ReturnValue;
//                cmd.Parameters.AddWithValue("@PROD", updNumPed.Value);
//                conn.Open();
//                cmd.ExecuteNonQuery();
//                int ret = Convert.ToInt32(cmd.Parameters["@PROD"].Value);
//                if (ret == 0) lblResposta.Text = "Nao existe ";
//                else lblResposta.Text = "Existe " + ret;
//                conn.Close();

//                Produtos();
//            }


//            private void btnExecParam_Click_1(object sender, EventArgs e)
//            {
//                SqlCommand cmd = conn.CreateCommand();
//                cmd.CommandType = CommandType.StoredProcedure;
//                cmd.CommandText = "STP_COPIA_PEDIDO_P";
//                cmd.Parameters.AddWithValue("@PROD", updNumPed.Value);

//                cmd.Parameters.Add("@SELECT", SqlDbType.Int);
//                cmd.Parameters["@SELECT"].Direction = ParameterDirection.Output;

//                cmd.Parameters.Add("@MSG", SqlDbType.VarChar, 1000);
//                cmd.Parameters["@MSG"].Direction = ParameterDirection.Output;

//                // para receber o RETURN da procedure
//                cmd.Parameters.Add("@RETURN_VALUE", SqlDbType.Int);
//                cmd.Parameters["@RETURN_VALUE"].Direction = ParameterDirection.ReturnValue;
//                // executar a procedure
//                try
//                {
//                    // executar a procedure
//                    conn.Open();
//                    cmd.ExecuteNonQuery();
//                    // ler os parâmetros de retorno
//                    int numPed = Convert.ToInt32(cmd.Parameters["@SELECT"].Value);
//                    string msg = cmd.Parameters["@MSG"].Value.ToString();
//                    int ret = Convert.ToInt32(cmd.Parameters["@RETURN_VALUE"].Value);
//                    if (ret > 0) lblResposta.Text = "Erro: " + msg;
//                    else lblResposta.Text = "Gerado pedido num " + ret + "  " + numPed;
//                }
//                catch (Exception ex)
//                {
//                    MessageBox.Show(ex.Message);
//                }
//                finally
//                {
//                    conn.Close();
//                }
//            }


//            private void btnExecReader_Click_1(object sender, EventArgs e)
//            {
//                // definir o comando que será executado
//                SqlCommand cmd = conn.CreateCommand();
//                cmd.CommandText = "exec STP_COPIA_PEDIDO_P " + updNumPed.Value;
//                // executar o comando
//                try
//                {
//                    conn.Open();
//                    // executar a procedure (devolve SELECT)
//                    SqlDataReader dr = cmd.ExecuteReader();
//                    // devolve uma única linha, não precisa de loop
//                    dr.Read();
//                    // ler as colunas
//                    int numPed = Convert.ToInt32(dr[0]);
//                    string msg = dr[1].ToString();
//                    // fechar o DataReader
//                    dr.Close();
//                    // mostrar resultado
//                    if (numPed < 0)
//                        lblResposta.Text = "Erro: " + msg;
//                    else
//                        lblResposta.Text = "Gerado pedido número " + numPed;
//                }
//                catch (Exception ex)
//                {
//                    MessageBox.Show(ex.Message);
//                }
//                finally
//                {
//                    conn.Close();
//                }
//            }



//            private void btnExecParam_Click(string prod)
//            {

//                // definir o comando
//                SqlCommand cmd = conn.CreateCommand();
//                // quando a procedure retorna parâmetros de OUTPUT não é
//                // possível montar o comando EXEC, temos que fazer o seguinte:
//                cmd.CommandType = CommandType.StoredProcedure;
//                cmd.CommandText = "SP_JACA_CARGA_PRODUTOS";
//                // passar parâmetro de INPUT, o único que tem valor antes
//                // da execução da procedure
//                cmd.Parameters.AddWithValue("@prod", prod);
//                // parâmetros de OUTPUT, só terão valor após a execução
//                //cmd.Parameters.Add("@NUM_PEDIDO_NOVO", SqlDbType.Int);
//                //cmd.Parameters["@NUM_PEDIDO_NOVO"].Direction =
//                //                                  ParameterDirection.Output;
//                // para receber o RETURN da procedure
//                cmd.Parameters.Add("@RETURN_VALUE", SqlDbType.Int);
//                cmd.Parameters["@RETURN_VALUE"].Direction = ParameterDirection.ReturnValue;
//                // executar a procedure
//                try
//                {
//                    conn.Open();
//                    cmd.ExecuteNonQuery();
//                    int numPed = Convert.ToInt32(
//                                      cmd.Parameters["@prod"].Value);
//                    int ret = Convert.ToInt32(
//                                      cmd.Parameters["@RETURN_VALUE"].Value);
//                }
//                catch (Exception ex)
//                {
//                    MessageBox.Show(ex.Message);
//                }
//                finally
//                {
//                    conn.Close();
//                }
//            }


//            public void Produtos()
//            {

//                string filePath = null;
//                string caminho = null;

//                int linha = 1;
//                int numRepetidos = 0, numCarregados = 0;

//                SqlConnection conn = new SqlConnection("Data Source=BRCAENRODRIGUES\\MSSQLSERVER01; Integrated Security=True; Initial Catalog=LAMPADA");
//                SqlConnection conn1 = new SqlConnection("Data Source=BRCAENRODRIGUES\\MSSQLSERVER01; Integrated Security=True; Initial Catalog=LAMPADA");

//                filePath = @"C:\Base\info\produtos\produto.xlsx";
//                caminho = filePath;

//                FileInfo existingFile = new FileInfo(filePath);
//                ExcelPackage package = new ExcelPackage(existingFile);
//                / ExcelWorksheet workSheet = package.Workbook.Worksheets[cmbPlanilha.SelectedIndex + 1];
//                ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
//                StringBuilder conteudo = new StringBuilder();
//                var lista = new List<String>();
//                SqlCommand cmd = conn.CreateCommand();
//                ArrayList Excel = new ArrayList();
//                ArrayList SQL = new ArrayList();
//                IEnumerable<object> distinctItemsExcel = null;
//                ArrayList repetido = new ArrayList();
//                ArrayList carregado = new ArrayList();

//                distinctItemsExcel = Excel.Cast<object>().Distinct();
//                int totalExcel = Excel.Cast<object>().Count();
//                int distinctCount = Excel.Cast<object>().Distinct().Count();

//                for (int o = workSheet.Dimension.Start.Row + 1; o <= workSheet.Dimension.End.Row; o++)
//                {
//                    Excel.Add(workSheet.Cells[o, 1].Value.ToString());
//                }

//                if (conn.State.ToString() == "Closed")
//                {
//                    conn.Open();
//                }


//                for (int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
//                {


//                    for (int j = workSheet.Dimension.Start.Column; j <= workSheet.Dimension.End.Column; j++)
//                    {


//                        if ((j == 1) && workSheet.Cells[i, j].Value != null)
//                        {

//                            SqlCommand cmdProc = conn.CreateCommand();
//                            cmdProc.CommandType = CommandType.StoredProcedure;
//                            cmdProc.CommandText = "[SP_JACA_PRODUTOS]";
//                            cmdProc.Parameters.Add("@PROD", SqlDbType.Int);
//                            cmdProc.Parameters["@PROD"].Direction = ParameterDirection.ReturnValue;
//                            cmdProc.Parameters.AddWithValue("@PROD", workSheet.Cells[i, j].Value);


//                            if (conn.State.ToString() == "Closed")
//                                conn.Open();

//                            cmdProc.ExecuteNonQuery();
//                            int ret = Convert.ToInt32(cmdProc.Parameters["@PROD"].Value);
//                            conn.Close();


//                            if (ret == 1)
//                            {
//                                numRepetidos++;
//                                lblRepetido.Text = numRepetidos.ToString();
//                                lblRepetido.Refresh();
//                            }
//                            else
//                            {
//                                numCarregados++;
//                                lblCarregada.Text = numCarregados.ToString();
//                                lblCarregada.Refresh();
//                            }



//                            conteudo.Append(" declare @pro_id varchar(max) = '" + workSheet.Cells[i, j].Value + "';");
//                            conteudo.Append(Environment.NewLine);
//                            conteudo.Append(" if  (select max(pro_id) from D_Produtos where pro_id = @pro_id) = (select (pro_id) from D_Produtos where pro_id = (@pro_id)) ");
//                            conteudo.Append(Environment.NewLine);
//                            conteudo.Append(" print 'OK' ");
//                            conteudo.Append(Environment.NewLine);
//                            conteudo.Append(" else ");
//                            conteudo.Append(" INSERT INTO D_PRODUTOS " +
//                                            "(PRO_ID, " +
//                                            "PRO_DESCRICAO, " +
//                                            "PRO_UND_ID, " +
//                                            "PRO_NCM, " +
//                                            "PRO_MARGEM, " +
//                                            "[Arq_Origem_ID]) " +
//                                            " VALUES ( ");
//                            conteudo.Append("'" + workSheet.Cells[i, j].Value + "', ");
//                        }
//                        else if (j == workSheet.Dimension.End.Column)
//                        {
//                            conteudo.Append(" 20, 1 ");
//                        }
//                        else
//                        {
//                            conteudo.Append(workSheet.Cells[i, j].Value == null ? " NULL " : "'" + workSheet.Cells[i, j].Value.ToString().Replace('\'', ' ') + "', ");
//                        }


//                    }

//                    if (i == workSheet.Dimension.End.Row)
//                    {
//                        conteudo.Append(" ) ");
//                    }
//                    else
//                    {
//                        conteudo.Append(")");
//                        conteudo.Append(Environment.NewLine);
//                    }

//                    if (conn.State.ToString() == "Closed")
//                    {
//                        conn.Open();
//                    }

//                    Clipboard.SetText(conteudo.ToString());
//                    linha = linha + 1;
//                    cmd.CommandText = conteudo.ToString();
//                    SqlTransaction trE = null;
//                    MessageBox.Show(conteudo.ToString());
//                    trE = conn.BeginTransaction();
//                    cmd.Transaction = trE;
//                    cmd.ExecuteNonQuery();
//                    trE.Commit();
//                    conteudo.Clear();

//                }
//                package.Dispose();

//            }
//                catch (Exception ex)
//                {
//                    MessageBox.Show(ex.Message);
//                }
//                finally
//                {

//                notification.ShowBalloonTip(5000);
//                   }
//    SqlCommand cmdArquivoCarregado = conn.CreateCommand();
//    cmdArquivoCarregado.CommandText =
//                " declare @tabela varchar(max) = 'D_Produtos';" +
//                " if (select count(arq_id) from S_ArquivoCarregado where Arq_Tabela = @tabela) = 0" +
//                " insert into S_ArquivoCarregado" +
//                " (Arq_ID, Arq_Nome, Arq_Tabela, Arq_Mensagem, Arq_DataCarga, Arq_Quantidade, Arq_Login)" +
//                " values(1, '" + caminho + "', @tabela, 'Carga efetuada com sucesso.'," +
//                " GETDATE(), ' " + linha.ToString() + "', REPLACE(SUSER_NAME(), 'ATRAME\\',''))" +
//                " else" +
//                " insert into S_ArquivoCarregado" +
//                " (Arq_ID, Arq_Nome, Arq_Tabela, Arq_Mensagem, Arq_DataCarga, Arq_Quantidade, Arq_Login)" +
//                " values(" + '1' + ", '" + caminho + "', @tabela, 'Carga efetuada com sucesso.'," +
//                " GETDATE(), " + linha.ToString() + ", REPLACE(SUSER_NAME(), 'ATRAME\\',''))";


//                if (conn.State.ToString() == "Closed")
//                {
//                    conn.Open();
//                }

//SqlTransaction trA = null;
//trA = conn.BeginTransaction();
//                cmdArquivoCarregado.Transaction = trA;
//                cmdArquivoCarregado.ExecuteNonQuery();
//                trA.Commit();
//                conn.Close();
//                conn1.Close();

//                MessageBox.Show(new Form { TopMost = true }, "Carregamento de " + numCarregados.ToString() + " registros de produtos realizados com sucesso");


//            }


//        }
//    }

}
}
