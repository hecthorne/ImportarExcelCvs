using ImportToXLS.Importacao;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ImportToXLS
{
    public static class ImportarCpfPlanilhaExcelCsv
    {
        public static void TratarArquivoExcel(int tamanhoConteudo, string extensaoArquivo, string nomeArquivo, Stream arquivoStream)
        {
            var resultado = ImportarExcel.ObterDadosDoArquivoExcel(tamanhoConteudo, extensaoArquivo, nomeArquivo, arquivoStream);

            var listaCpf = new List<string>();

            for (int i = 0; i < resultado.Tables[0].Rows.Count; i++)
            {
                listaCpf.Add(Regex.Replace(resultado.Tables[0].Rows[i][0].ToString(), @"\W+", ""));
            }
        }

        public static void TratarArquivoCsv(string nomeArquivo, Stream arquivoStream)
        {
            var resultado =  ImportarExcel.ObterDadosDoArquivoCsv(nomeArquivo, arquivoStream).Select(c => Regex.Replace(c, @"\W+", "")).ToList();            
        }
    }
}