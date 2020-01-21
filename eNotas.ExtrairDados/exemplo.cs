using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace eNotas.ExtrairDados
{
    class exemplo
    {
        /*
        protected void frmProgressExportarFatorConversao_DoWork(UI.Controles.Form.Progress sender, DoWorkEventArgs e)
        {
            //Force Globalization
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("pt-BR");

            //Argumento
            DataRow[] gridRow = (DataRow[])e.Argument;

            //Carrega Id's de Pedidos
            var ids = (from Row in gridRow
                       select Row["CPROD"].ToString()).ToList();

            //Busca Produtos
            DataTable dtProdutos = Negocio.Modelo.Produto.RelatorioFatorConversao(ids);

            try
            {
                #region Excel

                //Variáveis
                ExcelWorksheet worksheet = null;

                //Arquivo
                FileInfo fileInfo = new FileInfo(strCaminho);
                if (fileInfo.Exists)
                    fileInfo.Delete();

                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    if (dtProdutos.Rows.Count > 0)
                    {
                        //Total
                        int total = dtProdutos.Rows.Count;

                        //Variaveis Rows/Col
                        int ToRow = total + 2;
                        int ToCol = 5; //Total de colunas
                        int SumRow = ToRow + 1;

                        //Variaveis Progress
                        int count = 0;

                        // add a new worksheet to the empty workbook
                        worksheet = package.Workbook.Worksheets.Add("Produto");

                        #region Add the headers row 1

                        worksheet.Cells["A1"].Value = "Item";
                        worksheet.Cells["A1:C1"].Merge = true;
                        worksheet.Cells["D1"].Value = "Conversão";
                        worksheet.Cells["D1:E1"].Merge = true;

                        //Format row header 1 style;
                        using (var range = worksheet.Cells["A1:E1"])
                        {
                            range.Style.Font.Bold = true;
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            range.Style.Font.Color.SetColor(Color.Black);
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(176, 176, 176));
                            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        #endregion Add the headers row 1

                        #region Add the headers row 2

                        worksheet.Cells["A2"].Value = "Código";
                        worksheet.Cells["B2"].Value = "Descrição";
                        worksheet.Cells["C2"].Value = "Tipo do Item";
                        worksheet.Cells["D2"].Value = "Unidade de Conversão";
                        worksheet.Cells["E2"].Value = "Fator de Conversão";

                        //Format row header 1 style;
                        using (var range = worksheet.Cells["A2:E2"])
                        {
                            range.Style.Font.Bold = true;
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            range.Style.Font.Color.SetColor(Color.Black);
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(198, 198, 198));
                            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        #endregion Add the headers row 2

                        #region Add some items in the cells

                        List<string> lstCodigoReferenciado = new List<string>();

                        //Add some items in the cells...
                        int row = 2;
                        foreach (DataRow drProduto in dtProdutos.Rows)
                        {
                            //Check if cancellation requested
                            if (sender.CancellationPending)
                            {
                                e.Cancel = true;
                                return;
                            }

                            row++;

                            //Campos                                        
                            worksheet.SetValue(row, 1, drProduto["CODIGO_PRODUTO_REF"]);
                            worksheet.SetValue(row, 2, drProduto["DESCRICA_PRODUTO_REF"]);
                            worksheet.SetValue(row, 3, drProduto["TIPO_ITEM_PRODUTO_REF"].ToString().Trim());
                            worksheet.SetValue(row, 4, drProduto["UNID_CONV"]);
                            worksheet.SetValue(row, 5, drProduto["FAT_CONV"]);


                            //Adiciona codigos referenciados
                            lstCodigoReferenciado.Add(drProduto["CODIGO_PRODUTO_REF"].ToString());

                            //Increment count
                            count++;

                            //Porcentagem
                            int porcentagem = ((int)count * 100) / total;

                            //update the progress on UI
                            sender.SetProgress(porcentagem);
                        }

                        #region Merge Columns 1 and 2
                        // Get distinct elements and convert into a list again.
                        List<string> distinct = lstCodigoReferenciado.Distinct().ToList();
                        int rowMerge = 0;
                        row = 3;

                        for (int i = 0; i < distinct.Count; i++)
                        {
                            rowMerge = 0;

                            foreach (DataRow drProduto in dtProdutos.Rows)
                            {
                                if (distinct[i].ToString() == drProduto["CODIGO_PRODUTO_REF"].ToString())
                                {
                                    row++;
                                    rowMerge++;

                                    if (i % 2 == 0)
                                    {
                                        //Format rows style backcolor;
                                        using (var range = worksheet.Cells[row - 1, 1, row - 1, ToCol])
                                        {
                                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(216, 227, 242));
                                        }
                                    }
                                }
                            }

                            //Merge cells collumn 1 and 2
                            worksheet.Cells[row - rowMerge, 1, (row - 1), 1].Merge = true;
                            worksheet.Cells[row - rowMerge, 2, (row - 1), 2].Merge = true;
                        }

                        #endregion Merge Columns 1 and 2

                        #endregion Add some items in the cells

                        #region Format Column / Autofilter / Freeze line

                        //Format rows style border;
                        using (var range = worksheet.Cells[3, 1, ToRow, ToCol])
                        {
                            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        }

                        //Format the values collumn 1 and 2
                        using (var range = worksheet.Cells[3, 1, ToRow, 2])
                        {
                            range.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }

                        //Format the values
                        using (var range = worksheet.Cells[3, 3, ToRow, ToCol])
                        {
                            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }

                        //Create an autofilter for the range
                        worksheet.Cells[2, 1, ToRow, ToCol].AutoFilter = true;

                        //Excel line freeze
                        worksheet.View.FreezePanes(3, 1);

                        #endregion Format Column / Autofilter / Freeze line

                        #region Format type cells
                        //Format type cells
                        for (int i = 0; i < total; i++)
                        {
                            //Row
                            row = i + 3;

                            //Campos
                            worksheet.Cells[row, 1].Style.Numberformat.Format = "@";
                            worksheet.Cells[row, 2].Style.Numberformat.Format = "@";
                            worksheet.Cells[row, 3].Style.Numberformat.Format = "@";
                            worksheet.Cells[row, 4].Style.Numberformat.Format = "@";
                            worksheet.Cells[row, 5].Style.Numberformat.Format = "#,##0.000000";
                        }
                        #endregion Format type cells

                        //Autofit columns for all cells
                        worksheet.Cells.AutoFitColumns();
                    }

                    if (dtProdutos.Rows.Count == 0)
                        throw new ApplicationException("Itens selecionados não possuem fator de conversão");

                    // Change the sheet view to show it in page layout mode
                    worksheet.View.PageLayoutView = false;

                    // set some document properties
                    package.Workbook.Properties.Title = "Produto | Fator de Conversão";

                    // set some extended property values
                    package.Workbook.Properties.Company = "Infofisco Serviços de Informática Ltda";

                    // save our new workbook and we are done!
                    package.Save();
                }
                #endregion Excel
            }
            catch (ApplicationException ae)
            {
                alertas.Add(ae.Message);
            }
            catch (Exception ex)
            {
                erros.Add(ex.Message);
            }
        }
        */
    }
}
