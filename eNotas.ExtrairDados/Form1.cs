using HtmlAgilityPack;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace eNotas.ExtrairDados
{
    public partial class Form1 : Form
    {
        private enum Browser
        {
            Chrome,
            Edge,
            Firefox,
            InternetExplorer,
            Safari,
            Opera
        }

        public Form1()
        {
            InitializeComponent();
        }

        public string ConverterCaracteresEspeciais(string texto)
        {
            string tmp = texto;

            if (!string.IsNullOrEmpty(tmp))
            {
                tmp = tmp.Replace("&amp;", "&");
                tmp = tmp.Replace("&lt;", "<");
                tmp = tmp.Replace("&gt;", ">");
                tmp = tmp.Replace("&quot;", "\"");
                tmp = tmp.Replace("&#39;", "'");
            }

            return tmp;
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            //Variaveis
            IWebDriver driver = null;
            int ToRow = 0;
            int ToCol = 0;
            string munc = string.Empty;
            string uf = string.Empty;

            try
            {
                #region browser

                string url = @"https://e-gov.betha.com.br/e-nota/pesquisa_prestadores.faces";

                Browser browser = (Browser)Enum.Parse(typeof(Browser), ConfigurationManager.AppSettings["selenium_webdriver"].ToString());

                switch (browser)
                {
                    case Browser.Chrome:
                        {
                            //Create FireFox Service
                            OpenQA.Selenium.Chrome.ChromeDriverService chromeService = OpenQA.Selenium.Chrome.ChromeDriverService.CreateDefaultService();
                            chromeService.HideCommandPromptWindow = true;
                            chromeService.SuppressInitialDiagnosticInformation = true;
                            //Create FireFox Profile object
                            OpenQA.Selenium.Chrome.ChromeOptions chromeOptions = new OpenQA.Selenium.Chrome.ChromeOptions();
                            chromeOptions.AddArguments(string.Concat("--app=", url));
                            driver = new OpenQA.Selenium.Chrome.ChromeDriver(chromeService, chromeOptions);
                        }
                        break;
                    case Browser.Edge:
                        driver = new OpenQA.Selenium.Edge.EdgeDriver();
                        break;
                    case Browser.Firefox:
                        {
                            //Create FireFox Service
                            OpenQA.Selenium.Firefox.FirefoxDriverService firefoxService = OpenQA.Selenium.Firefox.FirefoxDriverService.CreateDefaultService(string.Concat(Application.StartupPath, @"\Selenium\Firefox"));
                            firefoxService.HideCommandPromptWindow = true;
                            firefoxService.SuppressInitialDiagnosticInformation = true;
                            //Create FireFox Profile object
                            OpenQA.Selenium.Firefox.FirefoxOptions firefoxOptions = new OpenQA.Selenium.Firefox.FirefoxOptions();
                            driver = new OpenQA.Selenium.Firefox.FirefoxDriver(firefoxService, firefoxOptions);
                            driver.Navigate().GoToUrl(url);
                        }
                        break;
                    case Browser.InternetExplorer:
                        {
                            //Create FireFox Service
                            OpenQA.Selenium.IE.InternetExplorerDriverService InternetExplorerService = OpenQA.Selenium.IE.InternetExplorerDriverService.CreateDefaultService(string.Concat(Application.StartupPath, @"\Selenium\InternetExplorerDriver"));
                            InternetExplorerService.HideCommandPromptWindow = true;
                            InternetExplorerService.SuppressInitialDiagnosticInformation = true;
                            //Create FireFox Profile object
                            OpenQA.Selenium.IE.InternetExplorerOptions InternetExplorerOptions = new OpenQA.Selenium.IE.InternetExplorerOptions();
                            driver = new OpenQA.Selenium.IE.InternetExplorerDriver(InternetExplorerService, InternetExplorerOptions);
                            driver.Navigate().GoToUrl(url);
                        }
                        break;
                    case Browser.Opera:
                        {
                            //Create FireFox Service
                            OpenQA.Selenium.Opera.OperaDriverService OperaService = OpenQA.Selenium.Opera.OperaDriverService.CreateDefaultService(string.Concat(Application.StartupPath, @"\Selenium\Opera"));
                            OperaService.HideCommandPromptWindow = true;
                            OperaService.SuppressInitialDiagnosticInformation = true;
                            //Create FireFox Profile object
                            OpenQA.Selenium.Opera.OperaOptions OperaOptions = new OpenQA.Selenium.Opera.OperaOptions();
                            driver = new OpenQA.Selenium.Opera.OperaDriver(OperaService, OperaOptions);
                            driver.Navigate().GoToUrl(url);
                        }
                        break;
                    default:
                        throw new NotSupportedException(string.Format(System.Globalization.CultureInfo.CurrentCulture, "Driver {0} não suportado", browser));
                }

                if (driver == null)
                    throw new Exception("WebDriver do Selenium não definido nas configurações");

                #endregion browser

                //Aguarda processamento da página 5min
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));

                //Arquivo
                FileInfo fileInfo = new FileInfo(string.Format(@"{0}\Prestadores.xls", Application.StartupPath));
                if (fileInfo.Exists)
                    fileInfo.Delete();

                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    try
                    {
                        #region UF

                        for (int i = 1; i < driver.FindElement(By.Id("mainForm:estado")).FindElements(By.TagName("option")).Count; i++)
                        {
                            try
                            {
                                if (i != 11)
                                    continue;

                                //Variaveis
                                ToRow = 1;
                                ToCol = 8; //Total de colunas
                                munc = string.Empty;
                                uf = string.Empty;

                                System.Threading.Thread.Sleep(1500);

                                //Seleciona UF
                                driver.FindElement(By.Id("mainForm:estado")).FindElements(By.TagName("option"))[i].Click();
                                uf = driver.FindElement(By.Id("mainForm:estado")).FindElements(By.TagName("option"))[i].Text;

                                System.Threading.Thread.Sleep(500);

                                //Worksheet
                                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(uf.Trim());

                                #region Municipio

                                for (int j = 1; j < driver.FindElement(By.Id("mainForm:municipio")).FindElements(By.TagName("option")).Count; j++)
                                {
                                    try
                                    {
                                        //if(j != 61)
                                        //    continue;

                                        //Municipio
                                        driver.FindElement(By.Id("mainForm:municipio")).FindElements(By.TagName("option"))[j].Click();
                                        munc = driver.FindElement(By.Id("mainForm:municipio")).FindElements(By.TagName("option"))[j].Text;

                                        if (!driver.FindElement(By.Id("mainForm:master:messageSection:warn")).Displayed)
                                        {
                                            string html = driver.PageSource;
                                            if (!string.IsNullOrEmpty(html))
                                            {
                                                // Load From String
                                                var htmlDocument = new HtmlAgilityPack.HtmlDocument();
                                                htmlDocument.LoadHtml(html);

                                                bool blnPrestador = htmlDocument.GetElementbyId("mainForm:prestadores") == null ? false : true;

                                                do
                                                {
                                                    htmlDocument.LoadHtml(driver.PageSource);
                                                    blnPrestador = htmlDocument.GetElementbyId("mainForm:prestadores") == null ? false : true;
                                                } while (blnPrestador == false);

                                                var prestadores = htmlDocument.GetElementbyId("mainForm:prestadores");
                                                var tabela = prestadores.SelectNodes("table");
                                                var linhas = tabela.ElementAt(0).SelectNodes("tbody//tr");

                                                #region Add the headers row 1

                                                worksheet.Cells["A1"].Value = "Razão/Nome";
                                                worksheet.Cells["B1"].Value = "Logradouro";
                                                worksheet.Cells["C1"].Value = "Bairro";
                                                worksheet.Cells["D1"].Value = "Complemento";
                                                worksheet.Cells["E1"].Value = "CEP";
                                                worksheet.Cells["F1"].Value = "Email";
                                                worksheet.Cells["G1"].Value = "Telefone";
                                                worksheet.Cells["H1"].Value = "UF";
                                                worksheet.Cells["I1"].Value = "Municipio";

                                                //Format row header 1 style;
                                                using (var range = worksheet.Cells["A1:I1"])
                                                {
                                                    range.Style.Font.Bold = true;
                                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                    range.Style.Font.Color.SetColor(Color.Black);
                                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(198, 198, 198));
                                                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                }

                                                #endregion Add the headers row 1

                                                #region Add some items in the cells

                                                foreach (HtmlNode linha in linhas)
                                                {
                                                    ToRow++;

                                                    foreach (HtmlNode campos in linha.SelectNodes("td"))
                                                    {
                                                        if (linha.SelectNodes("td").ElementAt(0).ChildNodes.ElementAt(0).HasChildNodes)
                                                        {
                                                            worksheet.Cells[ToRow, 1].Value = ConverterCaracteresEspeciais(linha.SelectNodes("td").ElementAt(0).ChildNodes.ElementAt(0).FirstChild.InnerText);
                                                            worksheet.Cells[ToRow, 2].Value = linha.SelectNodes("td").ElementAt(0).ChildNodes.ElementAt(3).InnerText.Replace("\n", "").Replace("\t", "").Replace("\r", "").Trim();
                                                            worksheet.Cells[ToRow, 3].Value = linha.SelectNodes("td").ElementAt(0).ChildNodes.ElementAt(5).InnerText.Replace("\n", "").Replace("\t", "").Replace("\r", "").Trim();
                                                            worksheet.Cells[ToRow, 4].Value = linha.SelectNodes("td").ElementAt(0).ChildNodes.ElementAt(7).InnerText.Replace("\n", "").Replace("\t", "").Replace("\r", "").Trim();
                                                            worksheet.Cells[ToRow, 5].Value = linha.SelectNodes("td").ElementAt(0).ChildNodes.ElementAt(9).InnerText.Replace("\n", "").Replace("\t", "").Replace("\r", "").Trim();
                                                        }

                                                        if (linha.SelectNodes("td").ElementAt(1).HasChildNodes)
                                                        {
                                                            worksheet.Cells[ToRow, 6].Value = linha.SelectNodes("td").ElementAt(1).ChildNodes.ElementAt(0).InnerText.Replace("\n", "").Replace("\t", "").Replace("\r", "").Trim();
                                                            worksheet.Cells[ToRow, 7].Value = linha.SelectNodes("td").ElementAt(1).ChildNodes.ElementAt(2).InnerText.Replace("\n", "").Replace("\t", "").Replace("\r", "").Trim();
                                                        }

                                                        string[] arr = uf.ToString().Split('-');

                                                        worksheet.Cells[ToRow, 8].Value = arr[0].Trim();
                                                        worksheet.Cells[ToRow, 9].Value = munc;
                                                    }
                                                }

                                                #endregion Add some items in the cells
                                            }

                                        }
                                    }
                                    catch(Exception ex) { throw; }
                                }

                                #endregion Municipio

                                #region Format type cells
                                //Format type cells
                                for (int fc = 0; fc < ToRow; fc++)
                                {
                                    //Row
                                    if (fc <= 1)
                                        continue;

                                    //Campos
                                    worksheet.Cells[fc, 1].Style.Numberformat.Format = "@";
                                    worksheet.Cells[fc, 2].Style.Numberformat.Format = "@";
                                    worksheet.Cells[fc, 3].Style.Numberformat.Format = "@";
                                    worksheet.Cells[fc, 4].Style.Numberformat.Format = "@";
                                    worksheet.Cells[fc, 5].Style.Numberformat.Format = "@";
                                    worksheet.Cells[fc, 6].Style.Numberformat.Format = "@";
                                    worksheet.Cells[fc, 7].Style.Numberformat.Format = "@";
                                    worksheet.Cells[fc, 8].Style.Numberformat.Format = "@";
                                }

                                #endregion Format type cells

                                if (ToRow == 1)
                                {
                                    worksheet = null;
                                    continue;
                                }

                                //Format the values
                                using (var range = worksheet.Cells[2, 1, ToRow, ToCol])
                                {
                                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                }

                                //Create an autofilter for the range
                                worksheet.Cells[1, 1, ToRow, ToCol].AutoFilter = true;

                                //Excel line freeze
                                worksheet.View.FreezePanes(2, 1);

                                //Autofit columns for all cells
                                worksheet.Cells.AutoFitColumns(0);

                                //// Change the sheet view to show it in page layout mode
                                //worksheet.View.PageLayoutView = false;
                            }
                            catch { }
                        }

                        #endregion UF
                    }
                    catch (Exception ex)
                    {
                        if (driver != null)
                            driver.Dispose();
                    }
                    finally
                    {
                        // set some document properties
                        package.Workbook.Properties.Title = "Prestadore e-Notas";

                        // save our new workbook and we are done!
                        package.Save();
                    }
                }
            }
            catch (Exception)
            {
                if (driver != null)
                    driver.Dispose();
            }
            finally
            {
                if (driver != null)
                    driver.Dispose();
            }
        }
    }
}
