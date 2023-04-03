using System;
using System.IO;
using OfficeOpenXml;
using System.Xml;
using System.Resources;
using System.Resources.NetStandard;

class Program
{
    static void Main(string[] args)
    {
        // define o contexto de licença
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // caminho para a pasta que contém os arquivos .xlsx
        string folderPath = @"C:\Users\joao.oliveira\Downloads\teste\TestexlsxtoResx\Nova pasta";

        // busca todos os arquivos .xlsx na pasta e suas subpastas
        string[] xlsxFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.AllDirectories);

        foreach (string filePath in xlsxFiles)
        {
            // obtém o caminho do diretório em que o arquivo .xlsx está localizado
            string fileDirectory = Path.GetDirectoryName(filePath);

            // cria um arquivo .resx com o mesmo nome do arquivo .xlsx
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string resxFilePath = Path.Combine(fileDirectory, $"{fileName}.resx");
            using ResXResourceWriter resxWriter = new ResXResourceWriter(resxFilePath);

            // lê o arquivo .xlsx
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                // itera sobre as linhas e salva as colunas 1 e 2 no arquivo .resx
                for (int row = 1; row <= rowCount; row++)
                {
                    string key = worksheet.Cells[row, 1].Value.ToString();
                    string value = worksheet.Cells[row, 2].Value.ToString();

                    resxWriter.AddResource(key, value);
                }
            }
        }
    }
}
