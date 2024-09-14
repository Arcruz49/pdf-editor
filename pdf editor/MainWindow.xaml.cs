using Microsoft.Win32;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf;
using System.Windows;
using pdf_editor.classes;
using System.Windows.Controls;
using System.Text.RegularExpressions;
using System.IO;
using SautinSoft;

namespace pdf_editor
{
    public partial class MainWindow : Window
    {

        public PdfDocument pdf;
        public string pagesStr;
        public PdfDocument pdfResultado;
        public MemoryStream pdfConvertido;
        public int fileType;

        public MainWindow()
        {
            InitializeComponent();
        }


        private void BtnAnexarPdf_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "PDF files (*.pdf)|*.pdf";

                if (openFileDialog.ShowDialog() == true)
                {
                    lblArquivoSelecionado.Content = openFileDialog.FileName;
                    lblArquivoSelecionadoConversao.Content = openFileDialog.FileName;
                    pdf = PdfReader.Open(openFileDialog.FileName, PdfDocumentOpenMode.Modify);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        #region remove páginas


        private void BtnRemoverPaginas_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string pagesStr = inputPaginas.Text;

                if (!ValidarPagesStr(pagesStr))
                    throw new Exception("Valores inválidos");

                if (pdf == null)
                    throw new Exception("Selecione um arquivo .pdf");

                if (string.IsNullOrEmpty(pagesStr))
                    throw new Exception("Selecione as páginas para remover");

                List<string> list = listaConfigPaginas(pagesStr);

                Retorno<PdfDocument> resultadoPdf = removePaginasPdf(list, pdf);

                if (!resultadoPdf.Success)
                    throw new Exception("Houve um erro ao remover as páginas: " + resultadoPdf.Message);

                pdfResultado = resultadoPdf.Dados.FirstOrDefault();

                if (pdfResultado == null)
                    throw new Exception("Erro ao gerar o PDF");

                btnDownloadPdf.Visibility = Visibility.Visible;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                btnDownloadPdf.Visibility = Visibility.Hidden;
            }
        }



        private List<string> listaConfigPaginas(string pagesStr)
        {
            pagesStr = pagesStr.Trim();
            string[] paginasArray = pagesStr.Split(',');


            return paginasArray.ToList();
        }

        private Retorno<PdfDocument> removePaginasPdf(List<string> paginas, PdfDocument pdf)
        {
            try
            {
                var paginasRevertidas = paginas.AsEnumerable().Reverse().ToList();
                foreach (var pag in paginasRevertidas)
                {
                    if (pag.Contains("-"))
                    {
                        string[] num = pag.Split('-');
                        var n1 = Convert.ToInt32(num[0]);
                        var n2 = Convert.ToInt32(num[1]);

                        if (n1 > n2) throw new Exception("Valores iniciais devem ser menores que os finais");


                        for (int i = n1; i <= n2; i++)
                        {
                            pdf.Pages.RemoveAt(Convert.ToInt32(n1) - 1);

                        }

                        

                    }
                    else
                    {
                        pdf.Pages.RemoveAt(Convert.ToInt32(pag) - 1);
                    }
                }


                return new Retorno<PdfDocument>
                {
                    Success = true,
                    Dados = { pdf }
                };
            }
            catch(Exception ex)
            {
                return new Retorno<PdfDocument>
                {
                    Success = false,
                    Message = ex.Message
                };
            }
        }

        private void BtnDownloadPdf_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "PDF files (*.pdf)|*.pdf";
                if (saveFileDialog.ShowDialog() == true)
                {
                    pdfResultado.Save(saveFileDialog.FileName);
                    MessageBox.Show("PDF salvo com sucesso!", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao salvar o PDF: " + ex.Message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool ValidarPagesStr(string pagesStr)
        {
            string pattern = @"^[\d,\s-]+$";

            return Regex.IsMatch(pagesStr, pattern);
        }


        #endregion

        #region conversão
        public void BtnConversao_Click(object sender, RoutedEventArgs e)
        {
            try
            {

                if (pdf == null)
                    throw new Exception("Selecione um pdf");

                ComboBoxItem selectedFormat = (ComboBoxItem)cbFormatos.SelectedItem;
                int selectedValue = Convert.ToInt32(selectedFormat.Tag);


                switch (selectedValue)
                {
                    case 0:
                            throw new Exception("Selecione um formato");

                    case 1:
                            var ret = convertePdfWord();
                            if (ret.Success == false)
                                throw new Exception(ret.Message);
                            break;
                }



                MessageBox.Show("PDF convertido com sucesso!");


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                btnDownloadPdfConvertido.Visibility = Visibility.Hidden;

            }
        }


        public void BtnDownloadPdfConvertido_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (pdfConvertido == null || pdfConvertido.Length == 0)
                    throw new Exception("Nenhum arquivo foi convertido.");

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                switch (fileType)
                {
                    case 1:
                        saveFileDialog.Filter = "Word Document (*.docx)|*.docx";
                        break;
                }
                
                saveFileDialog.Title = "Salvar Arquivo Convertido";

                if (saveFileDialog.ShowDialog() == true)
                {
                    using (FileStream fileStream = new FileStream(saveFileDialog.FileName, FileMode.Create, FileAccess.Write))
                    {
                        pdfConvertido.WriteTo(fileStream);
                    }

                    MessageBox.Show("Arquivo salvo com sucesso!", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao salvar o arquivo: " + ex.Message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public Retorno<MemoryStream> convertePdfWord()
        {
            try
            {
                PdfFocus pdfFocus = new PdfFocus();

                    pdfFocus.OpenPdf(lblArquivoSelecionado.Content.ToString());

                    pdfConvertido = new MemoryStream();

                    if (pdfFocus.ToWord(pdfConvertido) == 0)
                    {
                        pdfConvertido.Position = 0;

                        btnDownloadPdfConvertido.Visibility = Visibility.Visible;

                        fileType = 1;
                        
                        return new Retorno<MemoryStream>
                        {
                            Success = true,
                            Message = "PDF convertido para Word e armazenado na memória com sucesso!"
                        };
                    }
                return new Retorno<MemoryStream>
                {
                    Success = false,
                    Message = "Erro na conversão"
                };

            }
            catch (Exception ex)
            {
                return new Retorno<MemoryStream>
                {
                    Success = true,
                    Message = "Erro ao converter o PDF: " + ex.Message
                };
            }
        }



        //public Retorno<MemoryStream> convertePdfPowerPoint()
        //{
        //    try
        //    {
        //        PdfFocus pdfFocus = new PdfFocus();

        //        pdfFocus.OpenPdf(lblArquivoSelecionado.Content.ToString());

        //        pdfConvertido = new MemoryStream();

        //        if (pdfFocus.ToImage("teste") == 0)
        //        {
        //            pdfConvertido.Position = 0;

        //            btnDownloadPdfConvertido.Visibility = Visibility.Visible;

        //            fileType = 1;

        //            return new Retorno<MemoryStream>
        //            {
        //                Success = true,
        //                Message = "PDF convertido para Word e armazenado na memória com sucesso!"
        //            };
        //        }
        //        return new Retorno<MemoryStream>
        //        {
        //            Success = false,
        //            Message = "Erro na conversão"
        //        };

        //    }
        //    catch (Exception ex)
        //    {
        //        return new Retorno<MemoryStream>
        //        {
        //            Success = true,
        //            Message = "Erro ao converter o PDF: " + ex.Message
        //        };
        //    }
        //}


        #endregion

    }
}
