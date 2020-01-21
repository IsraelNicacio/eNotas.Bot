using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.PhantomJS;
using System.Windows.Forms;
using System.Configuration;

namespace Financeiro.Bots
{
    public static class Itau
    {
        /// <summary>
        /// Método sem recaptcha
        /// </summary>
        /// <param name="linhaDigitavel">Linha digitável do boleto</param>
        public static void Consultar(string cnpj, string id, string linhaDigitavel, string cpfCnpjPagador)
        {
            IWebDriver driver = null;

            try
            {
                //Trata linha digitável
                linhaDigitavel = linhaDigitavel.Replace(".", string.Empty).Replace(" ", string.Empty);

                #region Diretório / Arquivos

                DirectoryInfo directoryInfo = new System.IO.DirectoryInfo("Boletos");
                if (!directoryInfo.Exists)
                    directoryInfo.Create();

                directoryInfo = new System.IO.DirectoryInfo(Path.Combine(directoryInfo.FullName, cnpj));
                if (!directoryInfo.Exists)
                    directoryInfo.Create();

                directoryInfo = new System.IO.DirectoryInfo(Path.Combine(directoryInfo.FullName, string.Format("{0:yyyy-MM-dd}", DateTime.Now)));
                if (!directoryInfo.Exists)
                    directoryInfo.Create();

                FileInfo fileInfo = new System.IO.FileInfo(Path.Combine(directoryInfo.FullName, "Boletos.pdf"));
                if (fileInfo.Exists)
                    fileInfo.Delete();

                #endregion Diretório / Arquivo

                #region Chrome - Options

                if (ConfigurationManager.AppSettings["selenium_webdriver"] == "chrome")
                {
                    OpenQA.Selenium.Chrome.ChromeDriverService chromeService = OpenQA.Selenium.Chrome.ChromeDriverService.CreateDefaultService();
                    chromeService.HideCommandPromptWindow = true;
                    chromeService.SuppressInitialDiagnosticInformation = true;

                    OpenQA.Selenium.Chrome.ChromeOptions chromeOptions = new OpenQA.Selenium.Chrome.ChromeOptions();
                    chromeOptions.AddUserProfilePreference("download.default_directory", directoryInfo.FullName);
                    chromeOptions.AddUserProfilePreference("download.prompt_for_download", false);
                    chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");

                    //Disable
                    chromeOptions.AddArgument("disable-infobars");
                    //chromeOptions.AddArgument("headless");Utilizado para suprimir a exibição da janela do chrome

                    driver = new OpenQA.Selenium.Chrome.ChromeDriver(chromeService, chromeOptions);
                }

                #endregion Chrome - Options

                #region Firefox - Options

                if (ConfigurationManager.AppSettings["selenium_webdriver"] == "firefox")
                {
                    /*
                     * Firefoz config options
                     * 
                     * http://kb.mozillazine.org/About:config_entries#Browser.
                     * */

                    //Create FireFox Service
                    OpenQA.Selenium.Firefox.FirefoxDriverService firefoxService = OpenQA.Selenium.Firefox.FirefoxDriverService.CreateDefaultService();
                    firefoxService.HideCommandPromptWindow = true;
                    firefoxService.SuppressInitialDiagnosticInformation = true;

                    //Create FireFox Profile object
                    OpenQA.Selenium.Firefox.FirefoxOptions firefoxOptions = new OpenQA.Selenium.Firefox.FirefoxOptions();

                    //Set location to store files after downloading.
                    firefoxOptions.SetPreference("browser.download.folderList", 2);
                    firefoxOptions.SetPreference("browser.helperApps.alwaysAsk.force", false);
                    firefoxOptions.SetPreference("browser.download.manager.focusWhenStarting", false);
                    firefoxOptions.SetPreference("services.sync.prefs.sync.browser.download.manager.showWhenStarting", false);
                    firefoxOptions.SetPreference("pdfjs.disabled", true);
                    firefoxOptions.SetPreference("browser.download.dir", directoryInfo.FullName);

                    //Set Preference to not show file download confirmation dialogue using MIME types Of different file extension types.
                    firefoxOptions.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/pdf");

                    // Use this to disable Acrobat plugin for previewing PDFs in Firefox (if you have Adobe reader installed on your computer)
                    firefoxOptions.SetPreference("plugin.scan.Acrobat", "99.0");
                    firefoxOptions.SetPreference("plugin.scan.plid.all", false);

                    //Pass profile parameter In webdriver to use preferences to download file.
                    driver = new OpenQA.Selenium.Firefox.FirefoxDriver(firefoxService, firefoxOptions);
                }

                #endregion Firefox - Options

                if (driver == null)
                    throw new Exception("WebDriver do Selenium não definido nas configurações");

                driver.Navigate().GoToUrl("https://www.itau.com.br/servicos/boletos/atualizar/");

                //Aguarda processamento da página
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));

                //Textbox
                var elem = wait.Until(d =>
                {
                    try
                    {
                        var ele = d.FindElement(By.Id("representacaoNumerica"));
                        return ele.Displayed ? ele : null;
                    }
                    catch (UnhandledAlertException)
                    {
                        return null;
                    }
                    catch (StaleElementReferenceException)
                    {
                        return null;
                    }
                    catch (NoSuchElementException)
                    {
                        return null;
                    }
                });


                //Preenche os dados da pesquisa
                driver.FindElement(By.Id("representacaoNumerica")).SendKeys(linhaDigitavel);
                driver.FindElement(By.Id("txtDocumentoSacado")).SendKeys(cpfCnpjPagador);
                driver.FindElement(By.Id("btnProximo")).Click();

                //Aguarda
                System.Threading.Thread.Sleep(2500);

                //Acessa aba aerta
                driver.SwitchTo().Window(driver.WindowHandles.Last());

                //Botão de cownload
                elem = wait.Until(d =>
                {
                    try
                    {
                        var ele = d.FindElement(By.Name("frmPDF"));
                        return ele.Displayed ? ele : null;
                    }
                    catch (UnhandledAlertException)
                    {
                        return null;
                    }
                    catch (StaleElementReferenceException)
                    {
                        return null;
                    }
                    catch (NoSuchElementException)
                    {
                        return null;
                    }
                });

                //Download
                ((IJavaScriptExecutor)driver).ExecuteScript("javascript:document.frmPDF.submit();");

                //Aguarda
                System.Threading.Thread.Sleep(3000);

                //Renomear
                fileInfo.Refresh();
                if (fileInfo.Exists)
                    fileInfo.MoveTo(Path.Combine(directoryInfo.FullName, string.Format("{0}.pdf", id)));
            }
            catch
            {
                throw;
            }
            finally
            {
                if (driver != null)
                    driver.Dispose();
            }
        }
    }
}
