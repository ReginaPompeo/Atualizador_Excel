using OfficeOpenXml;
using System;
using System.IO;

namespace ManipulacaoPlanilha
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; //License

            string filePath = "G:\\CAMINHO\\DA\\PLANILHA\\Pasta1.xlsx";

            // Carregar o arquivo Excel usando EPPlus
            FileInfo fileInfo = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Definição de abas

                int rowCount = worksheet.Dimension.Rows;
                for (int i = 2; i <= rowCount; i++) // Definindo que começa na segunda linha, pois a primeira tem cabeçalho
                {
                    string cpf = worksheet.Cells[i, 1].Value?.ToString() ?? ""; // CPF está na primeira coluna
                    string nomeCompleto = worksheet.Cells[i, 2].Value?.ToString() ?? ""; // Nome completo está na segunda coluna

                    // Verificar se o nome completo está presente
                    if (!string.IsNullOrEmpty(nomeCompleto))
                    {
                        // Dividir o nome em nome e sobrenome
                        string[] partesNome = nomeCompleto.Split(' ');

                        string nome = partesNome[0];
                        string sobrenome = partesNome.Length > 1 ? string.Join(" ", partesNome, 1, partesNome.Length - 1) : "";

                        // Remoção de pontos, traços e barras do CPF
                        cpf = cpf.Replace(".", "").Replace("-", "").Replace("/", "");

                        // Atualizar células na planilha com os novos valores
                        worksheet.Cells[i, 1].Value = cpf;         
                        worksheet.Cells[i, 2].Value = nome;        
                        worksheet.Cells[i, 3].Value = sobrenome;
                    }
                }

                // Definição dos Cabeçalhos
                worksheet.Cells[1, 1].Value = "CPF";
                worksheet.Cells[1, 2].Value = "NOME";
                worksheet.Cells[1, 3].Value = "SOBRENOME";

                package.Save();
            }

            Console.WriteLine("Planilha atualizada com sucesso!");
        }
    }
}
