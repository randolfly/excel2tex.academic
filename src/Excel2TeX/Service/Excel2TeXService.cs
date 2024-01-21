using System.Data;
using System.Net;
using System.Text;
using Excel2TeX.Util;

namespace Excel2TeX.Service;

public class Excel2TeXService(ExcelIOService excelIOService)
{
    public ExcelIOService ExcelIOService { get; } = excelIOService;

    public async Task Excel2TeXAsync(List<string> filePathList)
    {
        ExportSettingFile(filePathList[0]);
        var taskArray = new Task[filePathList.Count];
        for (int i = 0; i < filePathList.Count; i++)
        {
            var dataTable = ExcelIOService.LoadExcelDataSet(filePathList[i]).Tables[0];
            var task = Excel2TeXAsync(filePathList[i]);
            taskArray[i] = task;
        }
        await Task.WhenAll(taskArray);
        Console.WriteLine("Excel Write Finished!");
    }

    public async Task Excel2TeXAsync(string filePath)
    {
        ExportSettingFile(filePath);
        var dataTable = ExcelIOService.LoadExcelDataSet(filePath).Tables[0];
        var outputFilePath = filePath.Replace(AppConfig.SourceFileSuffix,
            AppConfig.TargetFileSuffix);
        var texText = DataTable2TeX(dataTable);
        await File.WriteAllTextAsync(outputFilePath, texText);
    }

    public void Excel2TeX(List<string> filePathList)
    {
        ExportSettingFile(filePathList[0]);
        foreach (var filePath in filePathList)
        {
            Excel2TeX(filePath);
        }
        Console.WriteLine("Excel Write Finished!");
    }
    public void Excel2TeX(string filePath)
    {
        ExportSettingFile(filePath);
        var dataTable = ExcelIOService.LoadExcelDataSet(filePath).Tables[0];
        var outputFilePath = filePath.Replace(AppConfig.SourceFileSuffix,
            AppConfig.TargetFileSuffix);
        var texText = DataTable2TeX(dataTable);
        File.WriteAllText(outputFilePath, texText);
    }

    /// <summary>
    /// Export table setting file in the same path of excel file
    /// </summary>
    /// <param name="filePath">excel file path</param>
    private static void ExportSettingFile(string filePath)
    {
        var dirPath = Path.GetDirectoryName(filePath);
        var fullPath = Path.GetFullPath("setting.tex", dirPath);
        File.WriteAllText(fullPath, AppConfig.TeXSetting);
    }

    public static string DataTable2TeX(DataTable table)
    {
        var colCount = table.Columns.Count;
        var rowCount = table.Rows.Count;
        var sb = new StringBuilder();
        sb.AppendLine($"\\begin{{tabular}}{{*{{{colCount}}}{{c}}}}");
        // top line
        sb.InsertLine(colCount, 1.5f);
        // first row
        sb.InsertRow(table.Rows[0]);
        // second line
        sb.InsertLine(colCount, 0.75f);
        // data table content
        for (int i = 1; i < rowCount - 1; i++)
        {
            sb.InsertRow(table.Rows[i]);
            sb.InsertLine(colCount, null);
        }
        // last row
        sb.InsertRow(table.Rows[rowCount - 1]);
        // last line
        sb.InsertLine(colCount, 1.5f);
        sb.AppendLine(@"\end{tabular}");
        return sb.ToString();
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

public static class StringBuilderExtension
{
    public static void InsertLine(this StringBuilder sb, int colCount, float? lineWidth)
    {
        if (lineWidth is null)
        {
            sb.AppendLine(@"  \hhline{");
            for (int i = 0; i < colCount; i++)
            {
                sb.AppendLine($"    ~");
            }
            sb.AppendLine(@"  }");
        }
        else
        {
            sb.AppendLine(@"  \hhline{");
            for (int i = 0; i < colCount; i++)
            {
                sb.AppendLine($"  !{{\\sfill{{black}}{{{lineWidth:0.00}pt}}}}");
            }
            sb.AppendLine(@"  }");
        }

    }

    public static void InsertRow(this StringBuilder sb, DataRow dataRow)
    {
        sb.AppendLine($"  \\multicolumn{{1}}{{c}}{{{dataRow[0]}}}");
        for (int i = 1; i < dataRow.ItemArray.Length; i++)
        {
            sb.AppendLine($"   & \\multicolumn{{1}}{{c}}{{{dataRow[i]}}}");
        }
        sb.AppendLine(@"  \\");
    }

}