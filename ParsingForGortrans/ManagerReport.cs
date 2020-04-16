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

        private Dictionary<string, List<List<string>>> ReadExcelPage() 
        {
            string fileNameExcel = FileName;
            bool mark = false;
            bool emptyMark = false;
            var failInfo = new FileInfo(fileNameExcel);
            var dataList = new Dictionary<string,List<List<string>>>();
            using (var package = new ExcelPackage(failInfo))
            {
                var epWorkbook = package.Workbook;
                var worksheet = epWorkbook.Worksheets
                                          .First();
                var end = worksheet.Dimension.End;
                var bufferData = new List<List<string>>();
                for (int row = 1; row < end.Row; row++)
                {
                    var data = new List<string>();
                    for (int col = 1; col < end.Column; col++)
                    {
                        try
                        {
                            data.Add(worksheet.Cells[row, col]?.Text ?? "");
                        }
                        catch (ArgumentOutOfRangeException)
                        {
                            data.Add("");
                        }
                    }
                    if(data.Count(c=>!string.IsNullOrEmpty(c)) == 1) 
                    {
                        mark = true;
                    }
                    if(data.All(c => string.IsNullOrEmpty(c))) 
                    {
                        mark = false;
                        emptyMark = true;
                    }
                    if(mark && emptyMark)
                    {
                        string key = "-999";
                        try
                        {
                            key = bufferData.FirstOrDefault()
                                           ?.FirstOrDefault()
                                           ?.Split(' ')[2];
                        }
                        catch (ArgumentOutOfRangeException) 
                        {
                            key = Guid.NewGuid()
                                      .ToString();
                        }
                        dataList.Add(key, bufferData);
                        bufferData = new List<List<string>>
                        {
                            data
                        };
                        emptyMark = false;
                    }
                    else
                    {
                        bufferData.Add(data);
                    }
                }
            }
            return dataList;
        }

        private static RouteSheet SeparateStartData(List<List<string>> prototype)
        {
            var arrayDataBuffer = prototype?.FirstOrDefault()
                                           ?.FirstOrDefault(p => !string.IsNullOrEmpty(p))
                                           ?.Split(' ');
            var nullStringElement = prototype?.FirstOrDefault(l => l.All(a => string.IsNullOrEmpty(a)));
            var nullStringPosition = prototype?.IndexOf(nullStringElement);
            var fullNameRow = prototype.Count() > (nullStringPosition ?? 0) + 1
                            ? prototype.ElementAt((nullStringPosition ?? 0) + 1)
                            : new List<string> { "Ошибочный формат файла" };
            var fullName = fullNameRow.FirstOrDefault(f => f.Trim(' ').Length > 1);
            var shortName = arrayDataBuffer.Length > 1
                          ? arrayDataBuffer[0]
                          : "Ошибочный формат файла";
            return new RouteSheet(shortName, fullName);
        } 
        
        private static List<Pair> GetPairsForCrew(List<List<string>> data) 
        {
            var pairs = new List<Pair>();
            var activeData = data.GetRange(2, data.IndexOf(
                                              data.FirstOrDefault(f => f.All(a => string.IsNullOrEmpty(a)))) - 2);
            foreach (var item in activeData)
            {
                var pair = new Pair(item);
                pairs.Add(pair);
            }
            return pairs;
        }

        private static List<CheckPoint> GetCheckPoints(List<List<string>> data) 
        {
            var checkPoints = new List<CheckPoint>();
            var startPoint = data.IndexOf(data.FirstOrDefault(f => f.All(a => string.IsNullOrEmpty(a))));
            var activeData = data.GetRange(startPoint+2,data.Count - (2+startPoint+2));
            foreach (var item in activeData)
            {
                var name = item?.FirstOrDefault();
                var list = item.GetRange(1, item.Count - 1);
                foreach (var itemIn in list)
                {
                    var checkPoint = new CheckPoint(name, itemIn);
                    checkPoints.Add(checkPoint);
                }
            }
            checkPoints.Sort();
            return checkPoints;
        }

        private List<RouteSheet> GetRouteSheets()
        {
            var routeSheets = new List<RouteSheet>();
            var massivData = ReadExcelPage();
            var prototype = massivData?.FirstOrDefault()
                                       .Value;
            var routeSheet = SeparateStartData(prototype);
            var crews = new List<Crew>();
            var poinpoint = new List<List<CheckPoint>>();
            foreach(var list in massivData)
            {
                var crew = new Crew(int.TryParse(list.Key, out var number)
                                    ? (crews.Any(c => c.Number == number)
                                       ? crews.Select(s => s.Number).Max() + 1
                                       : number)
                                    : 999);
                crew.SetListPair(GetPairsForCrew(list.Value));
                crews.Add(crew);
                var checkPoints = GetCheckPoints(list.Value);
                poinpoint.Add(checkPoints);
            }
            routeSheet.InitCrews(crews);
            var x = poinpoint;
            routeSheets.Add(routeSheet);
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
