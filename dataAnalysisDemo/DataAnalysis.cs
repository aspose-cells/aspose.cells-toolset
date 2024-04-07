
namespace dataAnalysisDemo
{
    using Aspose.Cells;
    using Aspose.Cells.Charts;
    using Aspose.Cells.Pivot;
    using System;

    internal class DataAnalysis
    {
        private string _Path {  get; set; }

        private Workbook _Workbook { get; set; }

        internal DataAnalysis(string path)
        {
            _Path = path;
            if (File.Exists(path))
            {
                DateTime startDateLoading = DateTime.Now;
                Console.WriteLine(startDateLoading.ToLongTimeString());

                _Workbook = new Workbook(_Path, new TxtLoadOptions() { Separator = ';', LoadFilter = new LoadFilter(LoadDataFilterOptions.CellData), CheckExcelRestriction = false, ConvertNumericData = false, ConvertDateTimeData = false }); 
                DateTime endDateLoading = DateTime.Now;
                Cells cells = _Workbook.Worksheets[0].Cells;
                int maxColumnIndex = cells.MaxDataColumn;
                for (int rowIndex = 1  ; rowIndex <= cells.MaxDataRow; rowIndex++)
                {
                    Cell cell = cells[rowIndex, maxColumnIndex];
                    if (cell.Type == CellValueType.IsNull)
                    {
                        cell.PutValue(DateTime.Now.ToString("g"));
                    }
                }

                int sheetIndex = _Workbook.Worksheets.Add();
                Worksheet sheet = _Workbook.Worksheets[sheetIndex];
                int pivotIndex =  sheet.PivotTables.Add("=IssuesData240407!$A$1:$L$45422", "A3", "UserCount");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];
                //pivotTable.RowGrand = false;
                pivotTable.ColumnGrand = false;
                pivotTable.AddFieldToArea(PivotFieldType.Row, 1);
                pivotTable.AddFieldToArea(PivotFieldType.Data, 2);
             
                pivotTable.RefreshData();


                
                //string range = CellsHelper.CellIndexToName(pivotTable.DataBodyRange.StartRow, pivotTable.DataBodyRange.StartColumn) + ":" + CellsHelper.CellIndexToName(pivotTable.DataBodyRange.EndRow, pivotTable.DataBodyRange.EndColumn);
                string range = CellsHelper.CellIndexToName(pivotTable.DataBodyRange.StartRow, pivotTable.DataBodyRange.EndColumn) + ":" + CellsHelper.CellIndexToName(pivotTable.DataBodyRange.EndRow, pivotTable.DataBodyRange.EndColumn);
                sheet.Cells["E4"].PutValue("Min");
                sheet.Cells["F4"].SetFormula(string.Format("=MIN({0})", range),null);
                sheet.Cells["E5"].PutValue("Quarter");
                sheet.Cells["F5"].SetFormula(string.Format("=QUARTILE({0},1)", range), null);
                sheet.Cells["E6"].PutValue("Two-Quarter");
                sheet.Cells["F6"].SetFormula(string.Format("=QUARTILE({0},2)", range), null);
                sheet.Cells["E7"].PutValue("Three-Quarter");
                sheet.Cells["F7"].SetFormula(string.Format("=QUARTILE({0},3)", range), null);
                sheet.Cells["E8"].PutValue("Max");
                sheet.Cells["F8"].SetFormula(string.Format("=MAX({0})", range), null);

                int chartIndex = sheet.Charts.Add(ChartType.Line, 17, 5, 27, 30);
                Chart chart = sheet.Charts[chartIndex];
                chart.PivotSource = "Sheet2!" + pivotTable.Name;
                chart.Title.Text = "title";
                int count = chart.NSeries.Count;
                for (int i = 0; i < count; i++)
                {
                    chart.NSeries[i].DataLabels.ShowValue =true;
                    chart.NSeries[i].DataLabels.Position = LabelPositionType.Above;
                }
                //range = CellsHelper.CellIndexToName(pivotTable.DataBodyRange.StartRow, pivotTable.DataBodyRange.StartColumn) + ":" + CellsHelper.CellIndexToName(pivotTable.DataBodyRange.EndRow, pivotTable.DataBodyRange.StartColumn);

                //chart.NSeries.Add(range, true);
                //range = CellsHelper.CellIndexToName(pivotTable.DataBodyRange.StartRow, pivotTable.DataBodyRange.EndColumn) + ":" + CellsHelper.CellIndexToName(pivotTable.DataBodyRange.EndRow, pivotTable.DataBodyRange.EndColumn);
                //chart.NSeries.CategoryData = range;
                Console.WriteLine(sheet.Cells.MaxDataRow);
                Console.WriteLine(endDateLoading.ToLongTimeString());
                Console.WriteLine( endDateLoading - startDateLoading);
                _Workbook.Save(@"D:\PScripts\Output\IssuesData240407.xlsx");
            }


        }

        private void AddDataVal()
        {

        }
    }
}
