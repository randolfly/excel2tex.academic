namespace Excel2TeX.Test;

public class ExcelIOServiceTest
{
    [Fact]
    public void LoadExcelFileTest()
    {
        var ExcelIOService = new ExcelIOService();
        ExcelIOService.LoadExcelFile("./asset/sample.xlsx");
        Assert.Single(ExcelIOService.ExcelDataSet.Tables);
        Assert.Equal(5, ExcelIOService.ExcelDataSet.Tables[0].Rows.Count);
        Assert.Equal(4, ExcelIOService.ExcelDataSet.Tables[0].Columns.Count);
    }
}
