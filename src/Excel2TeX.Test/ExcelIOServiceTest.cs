using Excel2TeX.Service;
namespace Excel2TeX.Test;

public class ExcelIOServiceTest
{
    [Fact]
    public void LoadExcelFileTest()
    {
        var ExcelIOService = new ExcelIOService();
        ExcelIOService.LoadExcelDataSet("./asset/sample.xlsx");
        Assert.NotNull(ExcelIOService.ExcelDataSet);
        Assert.Single(ExcelIOService.ExcelDataSet.Tables);
        Assert.Equal(5, ExcelIOService.ExcelDataSet.Tables[0].Rows.Count);
        Assert.Equal(4, ExcelIOService.ExcelDataSet.Tables[0].Columns.Count);
    }
}
