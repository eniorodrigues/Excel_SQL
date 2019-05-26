using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.Common;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data.Sql;
using System.Runtime.InteropServices;
using System.Configuration;
using Microsoft.Office.Interop.Excel;
using ExcelIt = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Drawing;
using testeCampos;
using System.ComponentModel;
using System.Text.RegularExpressions;

namespace testeCampos
{
    class Fornecedores
    {
        System.Text.StringBuilder builder = new System.Text.StringBuilder();
        //builder.Append(" INSERT INTO D_FORNECEDORES (FOR_ID, FOR_NOME, FOR_PSS_ID, FOR_VINC, FOR_VINC_DT_INI, FOR_VINC_DT_FIM, FOR_CNPJ) " + Environment.NewLine + " VALUES ( ");
        //            for (int r = 2; r <= MyApp.Workbooks[2].Worksheets[1].UsedRange.Rows.Count; r++)
        //            {
        //                for (int k = 1; k <= MyApp.Workbooks[2].Worksheets[comboBox2.SelectedIndex + 1].UsedRange.Columns.Count; k++)
        //                {
        //                        if (k == MyApp.Workbooks[2].Worksheets[comboBox2.SelectedIndex + 1].UsedRange.Columns.Count)
        //                        {
        //                           builder.Append(MyApp.Workbooks[2].Worksheets[1].Cells[r, k].Value2 == null ? " NULL " : "'" + MyApp.Workbooks[2].Worksheets[1].Cells[r, k].Value2.ToString() + "' ");
        //                        }
        //                        else
        //                        {
        //                            builder.Append(MyApp.Workbooks[2].Worksheets[1].Cells[r, k].Value2 == null ? " NULL, " : "'" + MyApp.Workbooks[2].Worksheets[1].Cells[r, k].Value2.ToString() + "', ");
        //                        }
        //                }
        //                if (r == MyApp.Workbooks[2].Worksheets[1].UsedRange.Rows.Count)
        //                {
        //                    builder.Append(") ");
        //                }
        //                else
        //                {
        //                    builder.Append("),");
        //                    builder.Append(Environment.NewLine);
        //                    builder.Append(" (");
        //                }
        //            }

        //    Clipboard.SetText(builder.ToString());
        //    MessageBox.Show(builder.ToString());
            
        //    System.Data.DataTable table = ConvertListToDataTable(colunas);
        //    dataGridView2.DataSource = table;
        //    label5.Text = comboBox2.SelectedItem.ToString();
    }
}
