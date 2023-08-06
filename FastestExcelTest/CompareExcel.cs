using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using ClosedXML.Excel;
using LargeXlsx;
using MiniExcelLibs;

namespace FastestExcelTest;

[SimpleJob(RunStrategy.Monitoring, iterationCount: 10)]
[MemoryDiagnoser]
public class CompareExcel
{
    public List<ExcelDataDTO> GetTestDataByRow(int row)
    {
        var list = new List<ExcelDataDTO>();

        for (var i = 0; i < row; i++)
        {
            list
                .Add(new ExcelDataDTO
                {
                    C1 = (i * 1000 + 0).ToString(),
                    C2 = (i * 1000 + 1).ToString(),
                    C3 = (i * 1000 + 2).ToString(),
                    C4 = (i * 1000 + 3).ToString(),
                    C5 = (i * 1000 + 4).ToString(),
                    C6 = (i * 1000 + 5).ToString(),
                    C7 = (i * 1000 + 6).ToString(),
                    C8 = (i * 1000 + 7).ToString(),
                    C9 = (i * 1000 + 8).ToString(),
                    C10 = (i * 1000 + 9).ToString(),
                    C11 = (i * 1000 + 10).ToString(),
                    C12 = (i * 1000 + 11).ToString(),
                    C13 = (i * 1000 + 12).ToString(),
                    C14 = (i * 1000 + 13).ToString(),
                    C15 = (i * 1000 + 14).ToString(),
                    C16 = (i * 1000 + 15).ToString(),
                    C17 = (i * 1000 + 16).ToString(),
                    C18 = (i * 1000 + 17).ToString(),
                    C19 = (i * 1000 + 18).ToString(),
                    C20 = (i * 1000 + 19).ToString(),
                });
        }

        return list;
    }

    public IEnumerable<List<ExcelDataDTO>> GetTestData()
    {
        yield return GetTestDataByRow(10_000);
        yield return GetTestDataByRow(300_000);
        yield return GetTestDataByRow(1_000_000);
    }


    [ArgumentsSource(nameof(GetTestData))]
    [Benchmark]
    public void LargeXlsx(List<ExcelDataDTO> list)
    {
        var filename = Guid.NewGuid();
        using var stream = new FileStream($"{filename}.xlsx", FileMode.Create, FileAccess.Write);
        using var xlsxWriter = new XlsxWriter(stream);

        xlsxWriter.BeginWorksheet("Sheet1", 1, 1);
        xlsxWriter.BeginRow();
        foreach (var l in list)
        {
            xlsxWriter.BeginRow();
            xlsxWriter.Write(l.C1);
            xlsxWriter.Write(l.C2);
            xlsxWriter.Write(l.C3);
            xlsxWriter.Write(l.C4);
            xlsxWriter.Write(l.C5);
            xlsxWriter.Write(l.C6);
            xlsxWriter.Write(l.C7);
            xlsxWriter.Write(l.C8);
            xlsxWriter.Write(l.C9);
            xlsxWriter.Write(l.C10);
            xlsxWriter.Write(l.C11);
            xlsxWriter.Write(l.C12);
            xlsxWriter.Write(l.C13);
            xlsxWriter.Write(l.C14);
            xlsxWriter.Write(l.C15);
            xlsxWriter.Write(l.C16);
            xlsxWriter.Write(l.C17);
            xlsxWriter.Write(l.C18);
            xlsxWriter.Write(l.C19);
            xlsxWriter.Write(l.C20);
        }
    }

    [Benchmark]
    [ArgumentsSource(nameof(GetTestData))]
    public void MiniXlsx(List<ExcelDataDTO> list)
    {
        var path = Guid.NewGuid() + ".xlsx";

        // create 
        using var stream = File.Create(path);
        stream.SaveAs(list);
    }

    [Benchmark]
    [ArgumentsSource(nameof(GetTestData))]
    public void ClosedXML(List<ExcelDataDTO> list)
    {
        var path = Guid.NewGuid() + ".xlsx";

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Sample Sheet");

        for (var i = 0; i < list.Count; i++)
        {
            var row = list[i];
            worksheet.Cell(i + 1, 1).Value = row.C1;
            worksheet.Cell(i + 1, 2).Value = row.C2;
            worksheet.Cell(i + 1, 3).Value = row.C3;
            worksheet.Cell(i + 1, 4).Value = row.C4;
            worksheet.Cell(i + 1, 5).Value = row.C5;
            worksheet.Cell(i + 1, 6).Value = row.C6;
            worksheet.Cell(i + 1, 7).Value = row.C7;
            worksheet.Cell(i + 1, 8).Value = row.C8;
            worksheet.Cell(i + 1, 9).Value = row.C9;
            worksheet.Cell(i + 1, 10).Value = row.C10;
            worksheet.Cell(i + 1, 11).Value = row.C11;
            worksheet.Cell(i + 1, 12).Value = row.C12;
            worksheet.Cell(i + 1, 13).Value = row.C13;
            worksheet.Cell(i + 1, 14).Value = row.C14;
            worksheet.Cell(i + 1, 15).Value = row.C15;
            worksheet.Cell(i + 1, 16).Value = row.C16;
            worksheet.Cell(i + 1, 17).Value = row.C17;
            worksheet.Cell(i + 1, 18).Value = row.C18;
            worksheet.Cell(i + 1, 19).Value = row.C19;
            worksheet.Cell(i + 1, 20).Value = row.C20;
        }

        workbook.SaveAs(path);
    }
}