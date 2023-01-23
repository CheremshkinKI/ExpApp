using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System;

namespace HowToExcel
{
    public class MarketExcelGenerator
    {
        public byte[] Generate(MarketReport report)
        {
            var package = new ExcelPackage();

            var sheet = package.Workbook.Worksheets
                .Add("Market Report");
            //Console.Write("Введите название для таблицы: ");
            //sheet.Cells["B6"].Value = "Table:";
            //sheet.Cells[6, 3].Value = Console.ReadLine();

            //sheet.Cells["B3"].Value = "Location:";
            //sheet.Cells["C3"].Value = $"{report.Company.Address}, " +
            //                          $"{report.Company.City}, " +
            //                          $"{report.Company.Country}";
            //sheet.Cells["B4"].Value = "Sector:";
            //sheet.Cells["C4"].Value = report.Company.Sector;
            //sheet.Cells["B5"].Value = report.Company.Description;


            //материалы
            sheet.Cells[8, 2, 8, 10].LoadFromArrays(new object[][]{ new []{ "Type", "Brand", "Colour", "Size", "Count", "Available", "Price", "Seller", "Link"} });
            var row = 9;
            var column = 2;
            foreach (var item in report.Mater)
            {
                sheet.Cells[row, column].Value = item.Type;
                sheet.Cells[row, column + 1].Value = item.Brand;
                sheet.Cells[row, column + 2].Value = item.Colour;
                sheet.Cells[row, column + 3].Value = item.Size;
                sheet.Cells[row, column + 4].Value = item.Count;
                sheet.Cells[row, column + 5].Value = item.Available;
                sheet.Cells[row, column + 6].Value = item.Price;
                sheet.Cells[row, column + 7].Value = item.Seller;
                sheet.Cells[row, column + 8].Value = item.Link;
                row++;
            }

            sheet.Cells[1, 1, row, column + 2].AutoFitColumns();
            sheet.Column(2).Width = 14;
            sheet.Column(3).Width = 12;

            sheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            sheet.Cells[8, 3, 8 + report.Mater.Length, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

            sheet.Cells[8, 2, 8, 10].Style.Font.Bold = true;
            sheet.Cells["B2:C4"].Style.Font.Bold = true;
            
            sheet.Cells[8, 2, 8 + report.Mater.Length, 10].Style.Border.BorderAround(ExcelBorderStyle.Double);
            sheet.Cells[8, 2, 8, 10].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            sheet.Protection.IsProtected = false;
            return package.GetAsByteArray();
        }

        public byte[] Generate1(MarketReport report)
        {
            var package = new ExcelPackage();

            var sheet = package.Workbook.Worksheets
                .Add("Market Report");

            //оборудование
            sheet.Cells[1, 1, 1, 4].LoadFromArrays(new object[][] { new[] { "Name", "Price", "Unit", "Resource" } });
            var row1 = 2;
            var column1 = 1;
            foreach (var item in report.Equip)
            {
                sheet.Cells[row1, column1].Value = item.Name;
                sheet.Cells[row1, column1 + 1].Value = item.Price;
                sheet.Cells[row1, column1 + 2].Value = item.Unit;
                sheet.Cells[row1, column1 + 3].Value = item.Resource;
                row1++;
            }

            sheet.Cells[1, 1, row1, column1 + 2].AutoFitColumns();
            sheet.Column(2).Width = 14;
            sheet.Column(3).Width = 12;

            sheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            sheet.Cells[1, 2, 1 + report.Equip.Length, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

            sheet.Cells[1, 1, 1, 4].Style.Font.Bold = true;
            sheet.Cells["A1:D1"].Style.Font.Bold = true;

            sheet.Cells[1, 1, 1 + report.Equip.Length, 4].Style.Border.BorderAround(ExcelBorderStyle.Double);
            sheet.Cells[1, 1, 1, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            //var capitalizationChart = sheet.Drawings.AddChart("FindingsChart", OfficeOpenXml.Drawing.Chart.eChartType.Line);
            //capitalizationChart.Title.Text = "Capitalization";
            //capitalizationChart.SetPosition(7, 0, 5, 0);
            //capitalizationChart.SetSize(800, 400);
            //var capitalizationData = (ExcelChartSerie)(capitalizationChart.Series.Add(sheet.Cells["B9:B28"], sheet.Cells["D9:D28"]));
            //capitalizationData.Header = report.Company.Currency;

            sheet.Protection.IsProtected = false;
            return package.GetAsByteArray();
        }

        public byte[] GenerateEmpty1(MarketReport report)
        {
            var package = new ExcelPackage();

            var sheet = package.Workbook.Worksheets
                .Add("Market Report");

            //оборудование
            sheet.Cells[1, 1, 1, 4].LoadFromArrays(new object[][] { new[] { "Name", "Price", "Unit", "Resource" } });
            var row1 = 2;
            var column1 = 1;
            foreach (var item in report.Equip)
            {
                sheet.Cells[row1, column1].Value = item.Name;
                sheet.Cells[row1, column1 + 1].Value = item.Price;
                sheet.Cells[row1, column1 + 2].Value = item.Unit;
                sheet.Cells[row1, column1 + 3].Value = item.Resource;
                row1++;
            }

            sheet.Cells[1, 1, row1, column1 + 2].AutoFitColumns();
            sheet.Column(2).Width = 14;
            sheet.Column(3).Width = 12;

            sheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            sheet.Cells[1, 2, 1 + report.Equip.Length, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

            sheet.Cells[1, 1, 1, 4].Style.Font.Bold = true;
            sheet.Cells["A1:D1"].Style.Font.Bold = true;

            sheet.Cells[1, 1, 1 + report.Equip.Length, 4].Style.Border.BorderAround(ExcelBorderStyle.Double);
            sheet.Cells[1, 1, 1, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            sheet.Protection.IsProtected = false;
            return package.GetAsByteArray();
        }

        public byte[] GenerateNew()
        {
            var package = new ExcelPackage();

            var sheet = package.Workbook.Worksheets.Add("Market Report");



            sheet.Protection.IsProtected = false;
            return package.GetAsByteArray();
        }
    }
}