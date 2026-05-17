using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace TorontoDaycares.Exporters
{
    public record class ExcelExporter(string fileName) : IExporter
    {
        private string FileName { get; } = fileName;

        public async Task ExportAsync(Models.DaycareSearchResponse response)
        {
            if (!Directory.Exists(Path.GetDirectoryName(FileName)))
            {
                throw new DirectoryNotFoundException($"Directory not found: {Path.GetDirectoryName(FileName)}");
            }

            var items = response.TopPrograms
                .GroupBy(x => x.Program.ProgramType)
                .ToDictionary(g => g.Key, g => g.Select(x => (x.Daycare, x.Program)).ToList());

            if (items.Count == 0)
            {
                throw new InvalidOperationException("No programs to export.");
            }

            using IWorkbook workbook = new XSSFWorkbook();
            var creationHelper = workbook.GetCreationHelper();
            var headerStyle = workbook.CreateCellStyle();
            var headerFont = workbook.CreateFont();
            headerFont.IsBold = true;
            headerStyle.SetFont(headerFont);

            foreach (var programType in items)
            {
                var worksheet = workbook.CreateSheet(programType.Key.ToString());
                var headerRow = worksheet.CreateRow(0);

                CreateTextCell(headerRow, 0, "Name", headerStyle);
                CreateTextCell(headerRow, 1, "Rating", headerStyle);
                CreateTextCell(headerRow, 2, "Capacity", headerStyle);
                CreateTextCell(headerRow, 3, "Vacancy", headerStyle);
                CreateTextCell(headerRow, 4, "Address", headerStyle);
                CreateTextCell(headerRow, 5, "Url", headerStyle);

                var rowIndex = 1;
                foreach (var item in programType.Value)
                {
                    var row = worksheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(item.Daycare.Name);
                    row.CreateCell(1).SetCellValue(item.Program.Rating.Value);
                    row.CreateCell(2).SetCellValue(item.Program.Capacity);
                    if (item.Program.Vacancy.HasValue)
                    {
                        row.CreateCell(3).SetCellValue(item.Program.Vacancy.Value);
                    }
                    row.CreateCell(4).SetCellValue(item.Daycare.Address);

                    var hyperlinkCell = row.CreateCell(5);
                    hyperlinkCell.SetCellValue(item.Daycare.Uri.ToString());
                    var hyperlink = creationHelper.CreateHyperlink(HyperlinkType.Url);
                    hyperlink.Address = item.Daycare.Uri.ToString();
                    hyperlinkCell.Hyperlink = hyperlink;
                    rowIndex++;
                }

                for (var columnIndex = 0; columnIndex <= 5; columnIndex++)
                {
                    worksheet.AutoSizeColumn(columnIndex);
                }
            }

            await using var stream = File.Create(FileName);
            workbook.Write(stream);
        }

        private static void CreateTextCell(IRow row, int columnIndex, string value, ICellStyle style)
        {
            var cell = row.CreateCell(columnIndex);
            cell.SetCellValue(value);
            cell.CellStyle = style;
        }
    }
}
