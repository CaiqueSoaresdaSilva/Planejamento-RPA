using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Planejamento_RPA
{
    class Processos
    {

        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Range range;


        #region "Metodos Principais"

        public void Acompanhamento_BV(string dir)
        {

            // DIRETORIOS
            string diretorio = dir + @"\Acompanhamento Campanha BV PAS LOGADAS\";
            string procOk = diretorio + @"PROC_OK\";
            string procErro = diretorio + @"ERRO\";
            string LOG = diretorio + @"LOGS\";


            string nomeArq = "";
            string nomeArquivo = "";
            int count_erro = 0;
            int linha_erro = 0;


            // COLUNAS EXCEL
            string Data = "";
            string Quantidade = "";
            string[] arquivos;
            string layout = "";

            Email envioEmail = new Email();
            List<string> errors = new List<string>();
            StringBuilder mensagem = new StringBuilder();

            arquivos = Directory.GetFiles(diretorio, "*.xls");

            if(arquivos.Length >= 1)
            {
                try
                {
                    foreach (string arq in arquivos)
                    {


                        nomeArq = arq;
                        nomeArquivo = nomeArq.Substring(diretorio.Length, nomeArq.Length - diretorio.Length).Replace(@"\", "");

                        int rCnt;
                        int rw = 0;
                        int cl = 0;

                        xlApp = new Excel.Application();
                        xlWorkBook = xlApp.Workbooks.Open(nomeArq, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                        // RENOVAÇÃO
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);

                        range = xlWorkSheet.UsedRange;
                        rw = range.Rows.Count;
                        cl = range.Columns.Count;

                        if ((range.Cells[1, 2] as Excel.Range).Value != null)
                        {
                            Data = (range.Cells[1, 2] as Excel.Range).Value.ToString();
                            if (Data.ToUpper().Replace(" ", "").Replace("_", "") != "DATA") layout = Data;
                        }
                        if ((range.Cells[1, 3] as Excel.Range).Value != null)
                        {
                            Quantidade = (range.Cells[1, 3] as Excel.Range).Value.ToString();
                            if (Quantidade.ToUpper().Replace("_", "").Replace(" ", "") != "QUANTIDADE") layout = Quantidade;
                        }

                        if (layout == "")
                        {
                            // LIMPA A TABELA ANTES DE DAR OS INSERTS
                            clDALSQL obdal = new clDALSQL();
                            obdal.Ambiente = clDALSQL.AmbienteExecucao.BDE;
                            obdal.ExecutaQuery("TRUNCATE TABLE DW.dbo.RPA_205_PA", "DW");
                            obdal.Dispose();

                            for (rCnt = 2; rCnt <= rw; rCnt++) // LINHAS 
                            {

                                Data = "";
                                Quantidade = "";


                                if ((range.Cells[rCnt, 2] as Excel.Range).Value != null) Data = (range.Cells[rCnt, 2] as Excel.Range).Value.ToString();
                                if ((range.Cells[rCnt, 3] as Excel.Range).Value != null) Quantidade = (range.Cells[rCnt, 3] as Excel.Range).Value.ToString();

                                
                                if (Quantidade == null || Quantidade == "")
                                {
                                    break;
                                }

                                //TRATA AS VARIAVEIS

                                if (Data != null) Data = Data.Replace(" ", "").Replace("/", "-");
                                if (Quantidade != null) Quantidade = Quantidade.Replace(",", ".").Replace("(", "").Replace(")", "").Replace(" ", "").Replace("/", "-");

                                Data = Data.Substring(0, 10);

                                DateTime date = DateTime.Parse(Data);
                                string formattedDate = date.ToString("yyyy-MM-dd");

                                Console.WriteLine("Inserindo dados BV Renovação linha: " + rCnt + ", Data: " + Data + ", Quantidade: " + Quantidade);

                                Acompanhamentos_Campanhas_BV("2050001", Quantidade, formattedDate);

                            }

                        }
                        else
                        {
                            count_erro = count_erro + 1;
                        }

                        // LEADS
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                        range = xlWorkSheet.UsedRange;
                        rw = range.Rows.Count;
                        cl = range.Columns.Count;

                        if ((range.Cells[1, 2] as Excel.Range).Value != null)
                        {
                            Data = (range.Cells[1, 2] as Excel.Range).Value.ToString();
                            if (Data.ToUpper().Replace(" ", "").Replace("_", "") != "DATA") layout = Data;
                        }
                        if ((range.Cells[1, 3] as Excel.Range).Value != null)
                        {
                            Quantidade = (range.Cells[1, 3] as Excel.Range).Value.ToString();
                            if (Quantidade.ToUpper().Replace("_", "").Replace(" ", "") != "QUANTIDADE") layout = Quantidade;
                        }

                        if (layout == "")
                        {

                            for (rCnt = 2; rCnt <= rw; rCnt++) // LINHAS 
                            {

                                Data = "";
                                Quantidade = "";


                                if ((range.Cells[rCnt, 2] as Excel.Range).Value != null) Data = (range.Cells[rCnt, 2] as Excel.Range).Value.ToString();
                                if ((range.Cells[rCnt, 3] as Excel.Range).Value != null) Quantidade = (range.Cells[rCnt, 3] as Excel.Range).Value.ToString();


                                if (Quantidade == null || Quantidade == "")
                                {
                                    break;
                                }

                                //TRATA AS VARIAVEIS

                                if (Data != null) Data = Data.Replace(" ", "").Replace("/", "-");
                                if (Quantidade != null) Quantidade = Quantidade.Replace(",", ".").Replace("(", "").Replace(")", "").Replace(" ", "").Replace("/", "-");

                                Data = Data.Substring(0, 10);

                                DateTime date = DateTime.Parse(Data);
                                string formattedDate = date.ToString("yyyy-MM-dd");

                                Console.WriteLine("Inserindo dados BV Leads: " + rCnt + ", Data: " + Data + ", Quantidade: " + Quantidade);

                                Acompanhamentos_Campanhas_BV("2050002", Quantidade, formattedDate);

                            }
                        }
                        else
                        {
                            count_erro = count_erro + 1;
                        }   


                        //Verifica se houve ocorrencia de erros
                        if (count_erro != 0)
                        {
                            xlWorkBook.Close(true, null, null);
                            xlApp.Quit();

                            Marshal.ReleaseComObject(xlWorkSheet);
                            Marshal.ReleaseComObject(xlWorkBook);
                            Marshal.ReleaseComObject(xlApp);
                            File.Move(nomeArq, procErro + nomeArquivo);

                            if (layout != "")
                            {
                                errors.Add(" Layout invalido, Cabecalho do excel incorreto: " + layout);
                            }

                            mensagem.AppendLine("Erros no processamento do arquivo " + nomeArquivo + ": ");

                            mensagem.AppendLine("<ul>");
                            foreach (string erro in errors)
                            {
                                mensagem.AppendLine("<li> " + erro + "</li>");
                            }
                            mensagem.AppendLine("</ul>");

                            //envioEmail.EnviarEmail("Processo do arquivo " + nomeArquivo + " -  ERRO", mensagem.ToString());

                            //Arquivo de log com Erro
                            string nomeArquivoLog_erro = LOG + nomeArquivo + "ERRO_log.txt";
                            StreamWriter writer_erro = new StreamWriter(nomeArquivoLog_erro);
                            writer_erro.WriteLine(mensagem);
                            writer_erro.Close();
                        }
                        else
                        {


                            mensagem.AppendLine("Processamento do arquivo: " + nomeArquivo + " realizado com sucesso.");

                            xlWorkBook.Close(true, null, null);
                            xlApp.Quit();

                            Marshal.ReleaseComObject(xlWorkSheet);
                            Marshal.ReleaseComObject(xlWorkBook);
                            Marshal.ReleaseComObject(xlApp);
                            string extensao = System.IO.Path.GetExtension(nomeArq);
                            string nomeArquivoSemExtensao = nomeArquivo.Substring(0, nomeArquivo.Length - extensao.Length);
                            File.Move(nomeArq, procOk + nomeArquivoSemExtensao + "_" + DateTime.Today.ToString("yyyy-MM-dd") + extensao);

                            //envioEmail.EnviarEmail("Processo do Arquivo", mensagem.ToString());

                            //Arquivo de log
                            string nomeArquivoLog = LOG + nomeArquivo + "_log.txt";
                            StreamWriter writer = new StreamWriter(nomeArquivoLog);
                            writer.WriteLine(mensagem);
                            writer.Close();
                        }
                    }
                }
                catch (Exception ex)
                {
                    ex.Message.ToString();
                    //ARQUIVO DE LOG
                    string nomeArquvoLog = LOG + nomeArquivo + "_ERRO_Log.txt";
                    string erro = "Erro ao processar o arquivo. Exceção: " + ex.ToString();
                    StreamWriter writer = new StreamWriter(nomeArquvoLog);
                    writer.WriteLine(erro);
                    writer.Close();
                    throw;
                }
            }
           



        }
        public void WhatsApp_Resumo_BV(string dir)
        {

            // DIRETORIOS
            string diretorio = dir + @"\WhatsApp Resumo\";
            string procOk = diretorio + @"PROC_OK\";
            string procErro = diretorio + @"ERRO\";
            string LOG = diretorio + @"LOGS\";


            string nomeArq = "";
            string nomeArquivo = "";
            int count_erro = 0;
            int linha_erro = 0;


            // COLUNAS EXCEL
            string Data = "";
            string BaseFone = "";
            string BaseCPF = "";
            string Enviado = "";
            string Entregue = "";

            string[] arquivos;
            string layout = "";

            Email envioEmail = new Email();
            List<string> errors = new List<string>();
            StringBuilder mensagem = new StringBuilder();


            // Copiar arquivo do Planejamento para processar, somente se nao existir na pasta de processado ok
            try
            {
                if (File.Exists(procOk + "WHATS RESUMO_" + DateTime.Today.ToString("yyyy-MM-dd") + ".xlsx") == false)
                {
                    File.Copy(@"\\10.0.0.8\setores$\PLANEJAMENTO\OPERACOES_E_RESULTADOS\OPERACOES E RESULTADOS\Control Desk\BV\Whats\WHATS RESUMO.xlsx", diretorio + "WHATS RESUMO.xlsx");
                }
            }
            catch (Exception ex)
            {
                ex.Message.ToString();
                //ARQUIVO DE LOG
                string nomeArquvoLog = LOG + "WHATS RESUMO_ERRO_Log.txt";
                string erro = "Erro ao processar o arquivo. Exceção: " + ex.ToString();
                StreamWriter writer = new StreamWriter(nomeArquvoLog);
                writer.WriteLine(erro);
                writer.Close();
                throw;
            }

            // Comecar a ler arquivos do diretorio
            arquivos = Directory.GetFiles(diretorio, "*.xls");

            if (arquivos.Length >= 1)
            {
                try
                {
                    foreach (string arq in arquivos)
                    {

                        nomeArq = arq;
                        nomeArquivo = nomeArq.Substring(diretorio.Length, nomeArq.Length - diretorio.Length).Replace(@"\", "");

                        int rCnt;
                        int rw = 0;
                        int cl = 0;

                        xlApp = new Excel.Application();
                        xlWorkBook = xlApp.Workbooks.Open(nomeArq, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                        // RESUMO
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["Total"];

                        range = xlWorkSheet.UsedRange;
                        rw = range.Rows.Count;
                        cl = range.Columns.Count;

                        if ((range.Cells[1, 1] as Excel.Range).Value != null)
                        {
                            Data = (range.Cells[1, 1] as Excel.Range).Value.ToString();
                            if (Data.ToUpper().Replace(" ", "").Replace("_", "") != "DATA") layout = Data;
                        }
                        if ((range.Cells[1, 2] as Excel.Range).Value != null)
                        {
                            BaseFone = (range.Cells[1, 2] as Excel.Range).Value.ToString();
                            if (BaseFone.ToUpper().Replace("_", "").Replace(" ", "") != "BASEFONE") layout = BaseFone;
                        }
                        if ((range.Cells[1, 3] as Excel.Range).Value != null)
                        {
                            BaseCPF = (range.Cells[1, 3] as Excel.Range).Value.ToString();
                            if (BaseCPF.ToUpper().Replace("_", "").Replace(" ", "") != "BASECPF") layout = BaseCPF;
                        }
                        if ((range.Cells[1, 4] as Excel.Range).Value != null)
                        {
                            Enviado = (range.Cells[1, 4] as Excel.Range).Value.ToString();
                            if (Enviado.ToUpper().Replace("_", "").Replace(" ", "") != "ENVIADO") layout = Enviado;
                        }
                        if ((range.Cells[1, 5] as Excel.Range).Value != null)
                        {
                            Entregue = (range.Cells[1, 5] as Excel.Range).Value.ToString();
                            if (Entregue.ToUpper().Replace("_", "").Replace(" ", "") != "ENTREGUE") layout = Entregue;
                        }

                        if (layout == "")
                        {
                            // LIMPA A TABELA ANTES DE DAR OS INSERTS
                            clDALSQL obdal = new clDALSQL();
                            obdal.Ambiente = clDALSQL.AmbienteExecucao.BDE;
                            obdal.ExecutaQuery("TRUNCATE TABLE DW.dbo.RPA_205_WHATSAPP", "DW");

                            for (rCnt = 2; rCnt <= rw; rCnt++) // LINHAS 
                            {

                                Data = "";
                                BaseFone = "";
                                BaseCPF = "";
                                Enviado = "";
                                Entregue = "";


                                if ((range.Cells[rCnt, 1] as Excel.Range).Value != null) Data = (range.Cells[rCnt, 1] as Excel.Range).Value.ToString();
                                if ((range.Cells[rCnt, 2] as Excel.Range).Value != null) BaseFone = (range.Cells[rCnt, 2] as Excel.Range).Value.ToString();
                                if ((range.Cells[rCnt, 3] as Excel.Range).Value != null) BaseCPF = (range.Cells[rCnt, 3] as Excel.Range).Value.ToString();
                                if ((range.Cells[rCnt, 4] as Excel.Range).Value != null) Enviado = (range.Cells[rCnt, 4] as Excel.Range).Value.ToString();
                                if ((range.Cells[rCnt, 5] as Excel.Range).Value != null) Entregue = (range.Cells[rCnt, 5] as Excel.Range).Value.ToString();

                                /*
                                if (BaseFone == null || BaseFone == "")
                                {
                                    break;
                                }
                                */

                                //TRATA AS VARIAVEIS

                                if (Data != null) Data = Data.Replace(" ", "").Replace("/", "-");
                                if (BaseFone != null) BaseFone = BaseFone.Replace(",", ".").Replace("(", "").Replace(")", "").Replace(" ", "").Replace("/", "-");
                                if (BaseCPF != null) BaseCPF = BaseCPF.Replace(",", ".").Replace("(", "").Replace(")", "").Replace(" ", "").Replace("/", "-");
                                if (Enviado != null) Enviado = Enviado.Replace(",", ".").Replace("(", "").Replace(")", "").Replace(" ", "").Replace("/", "-");
                                if (Entregue != null) Entregue = Entregue.Replace(",", ".").Replace("(", "").Replace(")", "").Replace(" ", "").Replace("/", "-");

                                Data = Data.Substring(0, 10);

                                DateTime date = DateTime.Parse(Data);
                                string formattedDate = date.ToString("yyyy-MM-dd");

                                Console.WriteLine("Inserindo dados BV WhatsApp | Linha: " + rCnt + ", Data: " + Data + ", BaseFone: " + BaseFone + ", BaseCPF: " + BaseCPF + ", Enviado: " + Enviado + ", Entregue: " + Entregue);

                                obdal.ExecutaQuery("INSERT INTO DW.dbo.RPA_205_WHATSAPP VALUES('2050001','" + formattedDate + "','" + BaseFone  + "','" + BaseCPF + "','" + Enviado + "','" + Entregue + "')", "DW");

                            }
                            obdal.Dispose();

                        }
                        else
                        {
                            count_erro = count_erro + 1;
                        }


                        //Verifica se houve ocorrencia de erros
                        if (count_erro != 0)
                        {
                            xlWorkBook.Close(true, null, null);
                            xlApp.Quit();

                            Marshal.ReleaseComObject(xlWorkSheet);
                            Marshal.ReleaseComObject(xlWorkBook);
                            Marshal.ReleaseComObject(xlApp);
                            File.Move(nomeArq, procErro + nomeArquivo);

                            if (layout != "")
                            {
                                errors.Add(" Layout invalido, Cabecalho do excel incorreto: " + layout);
                            }

                            mensagem.AppendLine("Erros no processamento do arquivo " + nomeArquivo + ": ");

                            mensagem.AppendLine("<ul>");
                            foreach (string erro in errors)
                            {
                                mensagem.AppendLine("<li> " + erro + "</li>");
                            }
                            mensagem.AppendLine("</ul>");

                            //envioEmail.EnviarEmail("Processo do arquivo " + nomeArquivo + " -  ERRO", mensagem.ToString());

                            //Arquivo de log com Erro
                            string nomeArquivoLog_erro = LOG + nomeArquivo + "ERRO_log.txt";
                            StreamWriter writer_erro = new StreamWriter(nomeArquivoLog_erro);
                            writer_erro.WriteLine(mensagem);
                            writer_erro.Close();
                        }
                        else
                        {


                            mensagem.AppendLine("Processamento do arquivo: " + nomeArquivo + " realizado com sucesso.");

                            xlWorkBook.Close(true, null, null);
                            xlApp.Quit();

                            Marshal.ReleaseComObject(xlWorkSheet);
                            Marshal.ReleaseComObject(xlWorkBook);
                            Marshal.ReleaseComObject(xlApp);
                            string extensao = System.IO.Path.GetExtension(nomeArq);
                            string nomeArquivoSemExtensao = nomeArquivo.Substring(0, nomeArquivo.Length - extensao.Length);
                            File.Move(nomeArq, procOk + nomeArquivoSemExtensao + "_" + DateTime.Today.ToString("yyyy-MM-dd") + extensao);

                            //envioEmail.EnviarEmail("Processo do Arquivo", mensagem.ToString());

                            //Arquivo de log
                            string nomeArquivoLog = LOG + nomeArquivo + "_log.txt";
                            StreamWriter writer = new StreamWriter(nomeArquivoLog);
                            writer.WriteLine(mensagem);
                            writer.Close();
                        }
                    }
                }
                catch (Exception ex)
                {
                    ex.Message.ToString();
                    //ARQUIVO DE LOG
                    string nomeArquvoLog = LOG + nomeArquivo + "_ERRO_Log.txt";
                    string erro = "Erro ao processar o arquivo. Exceção: " + ex.ToString();
                    StreamWriter writer = new StreamWriter(nomeArquvoLog);
                    writer.WriteLine(erro);
                    writer.Close();
                    throw;
                }
            }




        }

        #endregion


        #region "Metodos Auxiliares"

        #endregion


        #region "Acesso as procs"

        public void Acompanhamentos_Campanhas_BV(string cod_cliente, string Quantidade, string Data)
        {

            clDALSQL clDAL = new clDALSQL();
            clDAL.Ambiente = clDALSQL.AmbienteExecucao.BDE;

            clDAL.ClearParameters();
            clDAL.AddParameters("COD_CLIENTE", cod_cliente, SqlDbType.VarChar, ParameterDirection.Input, 7);
            clDAL.AddParameters("QUANTIDADE", Quantidade, SqlDbType.VarChar, ParameterDirection.Input, 3);
            clDAL.AddParameters("DATA", Data, SqlDbType.VarChar, ParameterDirection.Input, 10);

            clDAL.ExecutaProcedure("dbo.ACOMPANHAMENTO_CAMPANHAS_BV", "DW");
        }


        #endregion
    }
}
