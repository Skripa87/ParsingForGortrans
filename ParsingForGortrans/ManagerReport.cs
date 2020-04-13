using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Windows.Controls;
using OfficeOpenXml;
using ClosedXML.Excel;

namespace ParsingForGortrans
{
    public class ManagerReport
    {
        public ManagerReport(string fileName)
        {
            FileName = fileName;            
        }

        private string FileName { get; set; }        

        private List<List<object>> ReadExcelPage() 
        {
            string fileNameExcel = FileName;
            var failInfo = new FileInfo(fileNameExcel);
            var dataList = new List<List<object>>();
            using (var package = new ExcelPackage(failInfo))
            {
                var epWorkbook = package.Workbook;
                var worksheet = epWorkbook.Worksheets
                                          .First();
                var end = worksheet.Dimension.End;
                for (int row = 1; row < end.Row; row++)
                {
                    var data = new List<object>();
                    for (int col = 1; col < end.Column; col++)
                    {
                        ExcelRange val;
                        //try
                        //{
                            val = worksheet.Cells[row, col];
                        //}
                        //catch (ArgumentOutOfRangeException ex)
                        //{
                        //    val = new ExcelCell();
                        //}
                        if (val == null || val.ToString()
                                              .ToUpperInvariant()
                                              .Contains("Ошибка")) continue;
                        data.Add(val);
                    }
                    dataList.Add(data);
                }
            }
            return dataList;
        }

        private List<RouteSheet> GetRouteSheets()
        {
            var routeSheets = new List<RouteSheet>();
            var massivData = ReadExcelPage();
            return routeSheets;
        }

        private void SetFormat<T>(IXLRange range, T value)
        {
            range.Merge();
            range.Style.Font.FontColor = XLColor.Black;
            range.Style.Font.FontSize = 10;
            range.Style.Font.FontName = "Arial Cyr";
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            range.Style.Font.Bold = true;
            range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            range.Style
                 .Border
                 .TopBorder = XLBorderStyleValues.Thin;
            range.Style
                 .Border
                 .RightBorder = XLBorderStyleValues.Thin;
            range.Style
                 .Border
                 .LeftBorder = XLBorderStyleValues.Thin;
            range.Style
                 .Border
                 .BottomBorder = XLBorderStyleValues.Thin;
            range.SetValue(value);
        }

        private void SetFormat<T>(IXLCell cell, T value)
        {
            cell.Style
                .Border
                .TopBorder = XLBorderStyleValues.Thin;
            cell.Style
                .Border
                .BottomBorder = XLBorderStyleValues.Thin;
            cell.Style
                .Border
                .LeftBorder = XLBorderStyleValues.Thin;
            cell.Style
                .Border
                .RightBorder = XLBorderStyleValues.Thin;
            cell.Style
                .Font
                .FontName = "Calibri";
            cell.SetValue(value);
        }
        
        private XLWorkbook CreateWorksheetForUserPersonalReport(XLWorkbook workbook)
        {
            var worksheet = workbook.AddWorksheet("");
            
            worksheet.Columns()
                     .AdjustToContents();
            worksheet.Rows()
                     .AdjustToContents();
            try
            {
                workbook.Save();
            }
            catch (Exception ex)
            {
                // ignored
            }
            return workbook;
        }        

        public void GetReport()
        {
            var routeSheets = GetRouteSheets();
        }
    }
}
