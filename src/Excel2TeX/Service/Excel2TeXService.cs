using System.Data;
using System.Net;

namespace Excel2TeX.Service;

public class Excel2TeXService(ExcelIOService excelIOService)
{
    public ExcelIOService ExcelIOService { get; } = excelIOService;

    public void Excel2TeX(List<string> filePathList)
    {
        foreach (var filePath in filePathList)
        {
            var dataTable = ExcelIOService.LoadExcelDataSet(filePath).Tables[0];
            PrintDataTable(dataTable);
        }
    }

    public void Excel2TeX(string filePath)
    {
        var dataTable = ExcelIOService.LoadExcelDataSet(filePath).Tables[0];
        PrintDataTable(dataTable);
    }

    public static void PrintDataTable(DataTable table)
    {
        // calculate max width of each column
        int[] columnWidths = new int[table.Columns.Count];
        for (int i = 0; i < table.Columns.Count; i++)
        {
            columnWidths[i] = table.Columns[i].ColumnName.Length;
            foreach (DataRow row in table.Rows)
            {
                int len = row[i].ToString().Length;
                if (len > columnWidths[i])
                {
                    columnWidths[i] = len;
                }
            }
        }
        // print table header (first row)
        for (int i = 0; i < table.Columns.Count; i++)
        {
            Console.Write(table.Rows[index: 0][i]
                .ToString()
                .PadRight(columnWidths[i]));
            Console.Write(" | ");
        }
        Console.WriteLine();
        Console.WriteLine(new string('-', columnWidths.Sum() + 3 * columnWidths.Length - 1));
        // print table data
        for (int i = 1; i < table.Rows.Count; i++)
        {
            for (int j = 0; j < table.Columns.Count; j++)
            {
                Console.Write(table.Rows[i][j]
                    .ToString()
                    .PadRight(columnWidths[j]));
                Console.Write(" | ");
            }
            Console.WriteLine();
        }
        Console.WriteLine();
    }
}
