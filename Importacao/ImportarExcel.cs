using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace ImportToXLS.Importacao
{
    public static class ImportarExcel
    {
        public static DataSet ObterDadosDoArquivoExcel(int tamanhoConteudo, string extensaoArquivo, string nomeArquivo, Stream arquivoStream)
        {
            var dataSet = new DataSet();

            if (tamanhoConteudo > 0)
            {
                var localizacaoArquivo = TratarCaminhoArquivoLocal(nomeArquivo, arquivoStream);

                ObterDadosDoArquivoXlsOuXlsx(extensaoArquivo, dataSet, localizacaoArquivo);
            }

            return dataSet;
        }

        public static IEnumerable<string> ObterDadosDoArquivoCsv(string nomeArquivo, Stream arquivoStream)
        {
            try
            {
                var localizacaoArquivo = TratarCaminhoArquivoLocal(nomeArquivo, arquivoStream);

                var reader = new StreamReader(arquivoStream);

                return File.ReadLines(localizacaoArquivo);
            }
            catch (Exception)
            {
                throw;
            }
            
        }

        private static DataSet ObterDadosDoArquivoXlsOuXlsx(string extensaoArquivo, DataSet dataSet, string localizacaoArquivo)
        {
            try
            {
                var excelConnectionString = string.Empty;

                if (extensaoArquivo == ".xls")
                {
                    excelConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + localizacaoArquivo + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                }
                else if (extensaoArquivo == ".xlsx")
                {
                    excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + localizacaoArquivo + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                }

                var excelConnection = new OleDbConnection(excelConnectionString);
                excelConnection.Open();

                var dataTable = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                var excelSheets = new String[dataTable.Rows.Count];
                var t = 0;

                //excel data saves in temp file here.
                foreach (DataRow row in dataTable.Rows)
                {
                    excelSheets[t] = row["TABLE_NAME"].ToString();
                    t++;
                }

                var query = string.Format("Select * from [{0}]", excelSheets[0]);

                using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(query, excelConnection))
                {
                    dataAdapter.Fill(dataSet);
                }

            }
            catch (Exception)
            {
                throw;
            }
            
            return dataSet;
        }

        private static string TratarCaminhoArquivoLocal(string nomeArquivo, Stream stream)
        {
            try
            {
                var localizacaoArquivo = Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory) + "\\Content\\" + nomeArquivo;

                if (File.Exists(localizacaoArquivo)) File.Delete(localizacaoArquivo);

                SaveStreamToFile(localizacaoArquivo, stream);

                return localizacaoArquivo;
            }
            catch (Exception)
            {
                throw;
            }
           
        }

        private static void SaveStreamToFile(string fileFullPath, Stream stream)
        {
            try
            {
                if (stream.Length == 0) return;

                using (FileStream fileStream = File.Create(fileFullPath, (int)stream.Length))
                {
                    // Fill the bytes[] array with the stream data
                    byte[] bytesInStream = new byte[stream.Length];
                    stream.Read(bytesInStream, 0, (int)bytesInStream.Length);

                    // Use FileStream object to write to the specified file
                    fileStream.Write(bytesInStream, 0, bytesInStream.Length);
                }
            }
            catch (Exception)
            {
                throw;
            }
            
        }
    }
}