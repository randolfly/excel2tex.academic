using System.Data;
using ExcelDataReader;

namespace Excel2TeX.Service;

public class ExcelIOService
{
    public DataSet? ExcelDataSet { get; set; }
    public ExcelIOService()
    {
        System.Text.Encoding.RegisterProvider(
            System.Text.CodePagesEncodingProvider.Instance);
    }

    public DataSet LoadExcelDataSet(string filePath)
    {
        using var stream = File.Open(filePath,
            FileMode.Open, FileAccess.Read);
        // Auto-detect format, supports:
        //  - Binary Excel files (2.0-2003 format; *.xls)
        //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
        using var reader = ExcelReaderFactory.CreateReader(stream);
        return ExcelDataSet = reader.AsDataSet();
    }
}
