using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.IO;
using System.Windows.Controls;
using OfficeOpenXml;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2013.Drawing.Chart;
using System.Windows;
using DocumentFormat.OpenXml.Office.ActiveX;

namespace ParsingForGortrans
{
    public class ManagerReport
    {
        public ManagerReport(List<string> fileNames, List<string> fileNamesWeekend)
        {
            FileNames = fileNames ?? new List<string>();
            FileNamesWeekend = fileNamesWeekend ?? new List<string>();
        }

        private List<string> FileNames { get; set; }        
        private List<string> FileNamesWeekend { get; set; }


        private Dictionary<string, List<List<string>>> ReadExcelPage(string fileName) 
        {
            var fileNameExcel = fileName;
            var mark = false;
            var emptyMark = false;
            var failInfo = new FileInfo(fileNameExcel);
            var dataList = new Dictionary<string,List<List<string>>>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(failInfo))
            {
                var epWorkbook = package.Workbook;
                var worksheet = epWorkbook.Worksheets
                                          .First();
                var end = worksheet.Dimension.End;
                var bufferData = new List<List<string>>();
                string key = "";
                List<string> data = null;
                for (int row = 1; row <= end.Row; row++)
                {
                    data = new List<string>();
                    for (int col = 1; col <= end.Column; col++)
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
                    if(data.All(c => string.IsNullOrEmpty(c)) || row == end.Row) 
                    {
                        mark = false;
                        emptyMark = true;
                    }
                    if(mark && emptyMark)
                    {
                        key = "-999";
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
                try
                {
                    try
                    {
                        bufferData.LastOrDefault()?.Add(worksheet.Cells[end.Row, end.Column]?.Text ?? "");
                    }
                    catch (ArgumentOutOfRangeException)
                    {
                        data.Add("");
                    }
                    key = bufferData.FirstOrDefault()
                                   ?.FirstOrDefault()
                                   ?.Split(' ')[2];
                    dataList.Add(key, bufferData);
                }
                catch (IndexOutOfRangeException ex)
                {
                    //MessageBox.Show("Пропущена ошибка");
                }
            }
            return dataList;
        }

        private static RouteSheet SeparateStartData(List<List<string>> prototype, bool isWeekend)
        {
            var arrayDataBuffer = prototype?.FirstOrDefault()
                                           ?.FirstOrDefault(p => !string.IsNullOrEmpty(p))
                                           ?.Split(' ');
            var nullStringElement = prototype?.FirstOrDefault(l => l.All(a => string.IsNullOrEmpty(a)));
            var nullStringPosition = prototype?.IndexOf(nullStringElement);
            var fullNameRow = prototype.Count > (nullStringPosition ?? 0) + 1
                            ? prototype.ElementAt((nullStringPosition ?? 0) + 1)
                            : new List<string> { "Ошибочный формат файла" };
            var fullName = fullNameRow.FirstOrDefault(f => f.Trim(' ').Length > 1);
            var shortName = arrayDataBuffer.Length > 1
                          ? arrayDataBuffer[0]
                          : "Ошибочный формат файла";
            return new RouteSheet(shortName, fullName, isWeekend);
        } 
        
        private static List<Pair> GetPairsForCrew(List<List<string>> data) 
        {
            var pairs = new List<Pair>();
            var flights = GetFlights(GetCheckPoints(data));
            var activeData = data.GetRange(2, data.IndexOf(
                                              data.FirstOrDefault(f => f.All(a => string.IsNullOrEmpty(a)))) - 2);
            foreach (var item in activeData)
            {
                var pair = new Pair(item);
                var pFlights = flights.FindAll(f =>f.CheckPoints
                                                    .Last()
                                                    .PitStopTimeStart <= pair.EndWorkTime);
                pair.SetFligths(pFlights);
                foreach (var pFlight in pFlights)
                {
                    flights.Remove(pFlight);
                }
                pairs.Add(pair);
            }
            return pairs;
        }
        
        private static List<List<string>> DistinctData(List<List<string>> data)
        {
            var result = new List<List<string>>();
            result.Add(data.FirstOrDefault());
            data.Remove(data.FirstOrDefault());
            foreach (var item in data)
            {
                if (!result.Any(r => r.All(a => item.IndexOf(a) != -1)))
                {
                    result.Add(item);
                }
            }
            return result;
        }

        private static List<CheckPoint> GetCheckPoints(List<List<string>> data) 
        {
            var checkPoints = new List<CheckPoint>();
            var startPoint = data.IndexOf(data.FirstOrDefault(f => f.All(a => string.IsNullOrEmpty(a))));
            var activePreData = data.GetRange(startPoint+2,data.Count - (2+startPoint+1));
            var activeData = DistinctData(activePreData);
            foreach (var item in activeData)
            {
                var name = item?.FirstOrDefault();
                var list = item.GetRange(1, item.Count - 1);
                list.RemoveAll(string.IsNullOrWhiteSpace);
                if (list.Any(a => a.Contains("\n")) || (activeData.IndexOf(item) == activeData.Count - 1))
                {
                    for (var i =0; i < list.Count; i++)
                    {
                        if (list[i].Contains("\n")) continue;
                        list[i] += "\n ";
                    }
                }
                foreach (var itemIn in list)
                {
                    var checkPoint = new CheckPoint(name, itemIn);
                    checkPoints.Add(checkPoint);
                }
            }
            checkPoints.Sort();
            CompositCheckPoint(checkPoints);
            return checkPoints;
        }

        private static List<CheckPoint> CompositCheckPoint(List<CheckPoint> checkPoints)
        {
            var endPoints = checkPoints.FindAll(s => s.IsEndpoint);
            if(!endPoints.Any()) return checkPoints;
            var endPointNames = endPoints.Select(s => s.Name
                                                       .Trim(' ')
                                                       .ToUpperInvariant())
                                         .Distinct();
            var activData = checkPoints.FindAll(c => endPointNames.Contains(c.Name
                                                                             .Trim(' ')
                                                                             .ToUpperInvariant()));
            checkPoints.RemoveAll(c => endPointNames.Contains(c.Name
                                                               .Trim(' ')
                                                               .ToUpperInvariant()));
            var index = 0;
            activData.Sort();
            var bufferData = new List<CheckPoint> {activData[index]};
            var endPointsComposit = new List<CheckPoint>();
            do
            {
                index++;
                if(!string.Equals(activData[index]?.Name
                                                 .Trim(' ')
                                                 .ToUpperInvariant(),
                                 bufferData.LastOrDefault()
                                           ?.Name
                                           .Trim(' ')
                                           .ToUpperInvariant(),new StringComparison())) 
                {
                    endPointsComposit.Add(CheckPoint.CreateEndPointFromCheckPointGroup(bufferData));
                    bufferData = new List<CheckPoint>();
                }
                bufferData.Add(activData[index]);
            } while (index < activData.Count-1);
            endPointsComposit.Add(CheckPoint.CreateEndPointFromCheckPointGroup(bufferData));
            checkPoints.AddRange(endPointsComposit);
            checkPoints.Sort();
            return checkPoints;
        }

        private static List<Flight> GetFlights(List<CheckPoint> checkPoints)
        {
            var flights = new List<Flight>();
            var number = 1;
            List<CheckPoint> flightPoints = null;
            Flight flight = null;
            foreach (var item in checkPoints)
            {
                if (item.IsEndpoint || flightPoints == null)
                {
                    if (flightPoints != null)
                    {
                        flightPoints.Add(item);
                        flight = new Flight(number);
                        flight.InitCheckPoints(flightPoints);
                        flights.Add(flight);                        
                    }
                    flightPoints = new List<CheckPoint>();
                }
                flightPoints.Add(item);
            }            
            return flights;
        }

        private List<RouteSheet> GetRouteSheets()
        {
            var routeSheets = new List<RouteSheet>();
            foreach (var fileName in FileNames)
            {
                var massivData = ReadExcelPage(fileName);
                var prototype = massivData?.FirstOrDefault()
                                           .Value;
                var routeSheet = SeparateStartData(prototype, false);
                var crews = new List<Crew>();
                foreach (var list in massivData)
                {
                    var crew = new Crew(int.TryParse(list.Key, out var number)
                        ? (crews.Any(c => c.Number == number)
                            ? crews.Select(s => s.Number).Max() + 1
                            : number)
                        : 999);
                    crew.SetListPair(GetPairsForCrew(list.Value));
                    crews.Add(crew);
                }
                routeSheet.InitCrews(crews);
                routeSheets.Add(routeSheet);
            }
            foreach (var fileName in FileNamesWeekend)
            {
                var massivData = ReadExcelPage(fileName);
                var prototype = massivData?.FirstOrDefault()
                    .Value;
                var routeSheet = SeparateStartData(prototype, true);
                var crews = new List<Crew>();
                foreach (var list in massivData)
                {
                    var crew = new Crew(int.TryParse(list.Key, out var number)
                        ? (crews.Any(c => c.Number == number)
                            ? crews.Select(s => s.Number).Max() + 1
                            : number)
                        : 999);
                    crew.SetListPair(GetPairsForCrew(list.Value));
                    crews.Add(crew);
                }
                routeSheet.InitCrews(crews);
                routeSheets.Add(routeSheet);
            }
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

        private void CreateWorkBook(string fileName, RouteSheet routeSheet, RouteSheet routeSheetWeekend)
        {
            using (var workBook = new XLWorkbook())
            {
                var worksheet = workBook.AddWorksheet($"{routeSheet.ShortName}");
                SetFormat(worksheet.Cell("A1"), "Начальная остановка");
                SetFormat(worksheet.Cell("B1"), "Конечная остановка");
                SetFormat(worksheet.Cell("C1"), "График выход");
                SetFormat(worksheet.Cell("D1"), "Смена");
                SetFormat(worksheet.Cell("E1"), "Время выхода");
                SetFormat(worksheet.Cell("F1"), "Время возвращения");
                SetFormat(worksheet.Cell("G1"), "отстой (мин)");
                SetFormat(worksheet.Cell("H1"), "линейный отстой (мин)");
                SetFormat(worksheet.Cell("I1"), "Тип рейса");
                SetFormat(worksheet.Cell("J1"), "Пн");
                SetFormat(worksheet.Cell("K1"), "Вт");
                SetFormat(worksheet.Cell("L1"), "Ср");
                SetFormat(worksheet.Cell("M1"), "Чт");
                SetFormat(worksheet.Cell("N1"), "Пт");
                SetFormat(worksheet.Cell("O1"), "Сб");
                SetFormat(worksheet.Cell("P1"), "Вс");
                var row = 2;
                if (routeSheet?.Crews != null)
                {
                    foreach (var crew in routeSheet.Crews)
                    {
                        var crewNumber = crew.Number;
                        foreach (var pair in crew.Pairs)
                        {
                            var pairNumber = pair.Number;
                            foreach (var flight in pair.Flights)
                            {
                                if (flight == null || !flight.CheckPoints.Any()) continue;
                                SetFormat(worksheet.Cell(row, 1), flight?.CheckPoints
                                                                            ?.FirstOrDefault()
                                                                            ?.Name ?? "");
                                SetFormat(worksheet.Cell(row, 2), flight?.CheckPoints
                                                                      ?.LastOrDefault()
                                                                      ?.Name ?? "");
                                SetFormat(worksheet.Cell(row, 3), crewNumber);
                                SetFormat(worksheet.Cell(row, 4), pairNumber);
                                SetFormat(worksheet.Cell(row, 5), flight?.CheckPoints
                                                                        ?.FirstOrDefault()
                                                                        ?.Time.ToString().Substring(0, 5));
                                var lastTime = flight?.CheckPoints
                                                     ?.LastOrDefault();
                                var returnTime = lastTime.PitStopTimeStart != TimeSpan.Zero
                                               ? lastTime?.PitStopTimeStart
                                               : lastTime?.Time;
                                SetFormat(worksheet.Cell(row, 6), returnTime.ToString().Substring(0, 5));
                                var minutes = ((flight.CheckPoints
                                                    .LastOrDefault()
                                                    ?.Time ?? TimeSpan.Zero) - (flight.CheckPoints
                                                                                              .LastOrDefault()
                                                                                              ?.PitStopTimeStart
                                                                                          ?? TimeSpan.Zero)).TotalMinutes;
                                SetFormat(worksheet.Cell(row, 7), minutes);
                                SetFormat(worksheet.Cell(row, 8), 0);
                                SetFormat(worksheet.Cell(row, 9), "рейс");
                                SetFormat(worksheet.Cell(row, 10), 1);
                                SetFormat(worksheet.Cell(row, 11), 1);
                                SetFormat(worksheet.Cell(row, 12), 1);
                                SetFormat(worksheet.Cell(row, 13), 1);
                                SetFormat(worksheet.Cell(row, 14), 1);
                                SetFormat(worksheet.Cell(row, 15), 0);
                                SetFormat(worksheet.Cell(row, 16), 0);
                                row++;
                            }
                        }
                    }
                }
                if (routeSheetWeekend?.Crews != null)
                {
                    foreach (var crew in routeSheetWeekend.Crews)
                    {
                        var crewNumber = crew.Number;
                        foreach (var pair in crew.Pairs)
                        {
                            var pairNumber = pair.Number;
                            foreach (var flight in pair.Flights)
                            {
                                if (flight == null || !flight.CheckPoints.Any()) continue;
                                SetFormat(worksheet.Cell(row, 1), flight?.CheckPoints
                                                                            ?.FirstOrDefault()
                                                                            ?.Name ?? "");
                                SetFormat(worksheet.Cell(row, 2), flight?.CheckPoints
                                                                      ?.LastOrDefault()
                                                                      ?.Name ?? "");
                                SetFormat(worksheet.Cell(row, 3), crewNumber);
                                SetFormat(worksheet.Cell(row, 4), pairNumber);
                                SetFormat(worksheet.Cell(row, 5), flight?.CheckPoints
                                                                        ?.FirstOrDefault()
                                                                        ?.Time.ToString().Substring(0, 5));
                                var lastTime = flight?.CheckPoints
                                                     ?.LastOrDefault();
                                var returnTime = lastTime.PitStopTimeStart != TimeSpan.Zero
                                               ? lastTime?.PitStopTimeStart
                                               : lastTime?.Time;
                                SetFormat(worksheet.Cell(row, 6), returnTime.ToString().Substring(0, 5));
                                var minutes = ((flight.CheckPoints
                                                    .LastOrDefault()
                                                    ?.Time ?? TimeSpan.Zero) - (flight.CheckPoints
                                                                                              .LastOrDefault()
                                                                                              ?.PitStopTimeStart
                                                                                          ?? TimeSpan.Zero)).TotalMinutes;
                                SetFormat(worksheet.Cell(row, 7), minutes);
                                SetFormat(worksheet.Cell(row, 8), 0);
                                SetFormat(worksheet.Cell(row, 9), "рейс");
                                SetFormat(worksheet.Cell(row, 10), 0);
                                SetFormat(worksheet.Cell(row, 11), 0);
                                SetFormat(worksheet.Cell(row, 12), 0);
                                SetFormat(worksheet.Cell(row, 13), 0);
                                SetFormat(worksheet.Cell(row, 14), 0);
                                SetFormat(worksheet.Cell(row, 15), 1);
                                SetFormat(worksheet.Cell(row, 16), 1);
                                row++;
                            }
                        }
                    }
                }
                worksheet.Columns().AdjustToContents();
                worksheet.Rows().AdjustToContents();
                try
                {
                    workBook.SaveAs(fileName);
                }
                catch(Exception ex) 
                {
                    MessageBox.Show(ex.Message,"Ошибка сохранения файла!");
                }
            }
        }

        private void CreateReport(List<RouteSheet> routeSheets)
        {
            if (routeSheets == null) return;
            var fileName = "";
            var routeSheetsWeekEnd = routeSheets.FindAll(r => r.IsWeekend);
            routeSheets.RemoveAll(r => r.IsWeekend);
            var routeSheetShotNames = routeSheets.Select(s => s.ShortName);
            foreach (var routeSheet in routeSheets)
            {
                fileName = $"route_{routeSheet.ShortName}.xlsx";
                var routeSheetWeekEnd = routeSheetsWeekEnd.Find(r =>
                    string.Equals(r.ShortName, routeSheet.ShortName, new StringComparison()));
                CreateWorkBook(fileName,routeSheet,routeSheetWeekEnd);
            }
        }        

        public void GetReport()
        {
            //if (DateTime.Now < new DateTime(2020, 5, 5))
            //{
                CreateReport(GetRouteSheets());
            //}
            //else return;
        }
    }
}
