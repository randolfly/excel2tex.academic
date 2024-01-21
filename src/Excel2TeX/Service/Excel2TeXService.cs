using System.Data;

namespace Excel2TeX.Service;

public class Excel2TeXService(ExcelIOService excelIOService)
{
    public ExcelIOService ExcelIOService { get; } = excelIOService;

    public static void PrintDataTable(DataTable table)
    {
        // 计算每个列的最大宽度
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
        // 打印表头
        for (int i = 0; i < table.Columns.Count; i++)
        {
            Console.Write(table.Columns[i].ColumnName.PadRight(columnWidths[i]));
            Console.Write(" | ");
        }
        Console.WriteLine();
        Console.WriteLine(new string('-', columnWidths.Sum() + 3 * columnWidths.Length - 1));
        // 打印数据行
        foreach (DataRow row in table.Rows)
        {
            for (int i = 0; i < table.Columns.Count; i++)
            {
                Console.Write(row[i]
                        .ToString()
                        .PadRight(columnWidths[i]));
                Console.Write(" | ");
            }
            Console.WriteLine();
        }
    }
}
