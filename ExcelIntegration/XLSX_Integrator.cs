using ClosedXML.Excel;

public class XLSX_Integrator
{
    public void CreateFile(List<Carro> carros)
    {
        List<string> columns = new List<string> {"A","B","C","D"};
        using var wbook = new XLWorkbook();
        var ws = wbook.Worksheets.Add("Sheet1");
       
        ws.Cell("A1").Value = "Marca";
        ws.Cell("B1").Value = "Modelo";
        ws.Cell("C1").Value = "Ano";
       
       for (int i = 0; i < carros.Count; i++)
       {
          ws.Cell($"A{i+2}").Value = carros[i].Marca;
          ws.Cell($"B{i+2}").Value = carros[i].Modelo;
          ws.Cell($"C{i+2}").Value = carros[i].Ano;
       }
       
        wbook.SaveAs("carros-simple.xlsx");
    }
    public void ReadFile()
    {
        var xls = new XLWorkbook("/Users/eduardo/projetos/dotnet/excel-integration/ExcelIntegration/carros-simple.xlsx");
        var planilha = xls.Worksheets.First(w => w.Name == "Sheet1");
        var totalLinhas = planilha.Rows().Count();
        
        // primeira linha é o cabecalho
        for (int l = 2; l <= totalLinhas; l++)
        {
            var marca = planilha.Cell($"A{l}").GetString();
            var modelo  = planilha.Cell($"B{l}").GetString();
            var ano  = planilha.Cell($"C{l}").Value;
            Console.WriteLine($"{marca} - {modelo} - {ano}");
        }

        xls.Save();
    }
}