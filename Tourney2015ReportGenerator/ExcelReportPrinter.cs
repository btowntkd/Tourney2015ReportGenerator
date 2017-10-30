using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace Tourney2015ReportGenerator
{
    public class ExcelReportPrinter
    {
        private const int HeaderFontSize = 18;
        private const int GroupFontSize = 14;
        private const string GroupBackgroundColor = "";
        private readonly RegistrationReporter _reporter;
        public ExcelReportPrinter(RegistrationReporter reporter)
        {
            _reporter = reporter;
        }

        public void CreateReport(string outFilePath)
        {
            using (var excelApp = new NetOffice.ExcelApi.Application())
            {
                excelApp.DisplayAlerts = false;
                var workbook = excelApp.Workbooks.Add();

                CreateCompetitorListWorksheet(workbook.Worksheets[1] as NetOffice.ExcelApi.Worksheet);
                CreateSpectatorListWorksheet(workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[workbook.Worksheets.Count]) as NetOffice.ExcelApi.Worksheet);
                CreateRefereeListWorksheet(workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[workbook.Worksheets.Count]) as NetOffice.ExcelApi.Worksheet);
                CreateCoachListWorksheet(workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[workbook.Worksheets.Count]) as NetOffice.ExcelApi.Worksheet);
                CreateSparringListWorksheet(workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[workbook.Worksheets.Count]) as NetOffice.ExcelApi.Worksheet);
                //CreateFormsListWorksheet(workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[workbook.Worksheets.Count]) as NetOffice.ExcelApi.Worksheet);

                workbook.SaveAs(outFilePath);
                excelApp.Quit();
            }
        }

        protected void CreateCompetitorListWorksheet(NetOffice.ExcelApi.Worksheet worksheet)
        {
            var competitors = _reporter.GetCompetitors()
                .OrderBy(x => x.LastName);

            var row = 1;
            worksheet.Name = "Competitor List";
            PrintWorksheetTitle("Competitor List", worksheet, row++);

            var startRow = row;
            PrintTableColumnNames(worksheet, Competitor.TableFullColumnNames(), row++);
            foreach (var competitor in competitors)
            {
                PrintTableColumnData(worksheet, competitor.TableFullRowData(), row++);
            }
            var endRow = row - 1;
            string firstCol = ExcelColumnFromNumber(1);
            string lastCol = ExcelColumnFromNumber(Competitor.TableFullColumnNames().Length);
            worksheet.Columns[firstCol + ":" + lastCol].AutoFit();

            var tableRange = worksheet.Range(
                worksheet.Cells[startRow, firstCol],
                worksheet.Cells[endRow, lastCol]);

            var table = worksheet.ListObjects.Add(
                NetOffice.ExcelApi.Enums.XlListObjectSourceType.xlSrcRange,
                tableRange,
                Type.Missing,
                NetOffice.ExcelApi.Enums.XlYesNoGuess.xlYes,
                Type.Missing);
            table.Name = "CompetitorList";
        }

        protected void CreateSpectatorListWorksheet(NetOffice.ExcelApi.Worksheet worksheet)
        {
            var spectators = _reporter.GetSpectators()
                .OrderBy(x => x.LastName);

            var row = 1;
            worksheet.Name = "Spectator List";
            PrintWorksheetTitle("Spectator List", worksheet, row++);

            PrintTableColumnNames(worksheet, Spectator.TableColumnNames(), row++);
            foreach (var spectator in spectators)
            {
                PrintTableColumnData(worksheet, spectator.TableRowData(), row++);
            }
            string firstCol = ExcelColumnFromNumber(1);
            string lastCol = ExcelColumnFromNumber(Spectator.TableColumnNames().Length);
            worksheet.Columns[firstCol + ":" + lastCol].AutoFit();
        }

        protected void CreateRefereeListWorksheet(NetOffice.ExcelApi.Worksheet worksheet)
        {
            var referees = _reporter.GetReferees()
                .OrderBy(x => x.LastName);

            var row = 1;
            worksheet.Name = "Referee List";
            PrintWorksheetTitle("Referee List", worksheet, row++);

            PrintTableColumnNames(worksheet, Referee.TableColumnNames(), row++);
            foreach (var referee in referees)
            {
                PrintTableColumnData(worksheet, referee.TableRowData(), row++);
            }
            string firstCol = ExcelColumnFromNumber(1);
            string lastCol = ExcelColumnFromNumber(Referee.TableColumnNames().Length);
            worksheet.Columns[firstCol + ":" + lastCol].AutoFit();
        }

        protected void CreateCoachListWorksheet(NetOffice.ExcelApi.Worksheet worksheet)
        {
            var coaches = _reporter.GetCoaches()
                .OrderBy(x => x.SchoolName)
                .ThenBy(x => x.LastName);

            var row = 1;
            worksheet.Name = "Coach List";
            PrintWorksheetTitle("Coach List", worksheet, row++);

            PrintTableColumnNames(worksheet, Coach.TableColumnNames(), row++);
            foreach (var coach in coaches)
            {
                PrintTableColumnData(worksheet, coach.TableRowData(), row++);
            }
            string firstCol = ExcelColumnFromNumber(1);
            string lastCol = ExcelColumnFromNumber(Coach.TableColumnNames().Length);
            worksheet.Columns[firstCol + ":" + lastCol].AutoFit();
        }

        protected void CreateSparringListWorksheet(NetOffice.ExcelApi.Worksheet worksheet)
        {
            var row = 1;
            worksheet.Name = "Sparring Divisions";
            PrintWorksheetTitle("Sparring Divisions", worksheet, row++);

            var sparringCompetitors = _reporter.GetCompetitors()
                .Where(x => x.IsSparring);

            var ageDivisions = EventInfo.AgeDivisions;
            var ranks = EventInfo.Ranks;

            foreach (var age in ageDivisions)
            {
                foreach (var rank in ranks)
                {
                    var matchingCompetitorsByGender = sparringCompetitors
                        .Where(x => x.IsIncludedInAgeDivision(age) && x.IsIncludedInRank(rank))
                        .GroupBy(x => x.Gender);

                    foreach (var genderGroup in matchingCompetitorsByGender)
                    {
                        var competitors = genderGroup.OrderBy(x => x.Weight);

                        PrintGroupHeader(genderGroup.Key + ", " + age.Name + ", " + rank, worksheet, row++);
                        PrintTableColumnNames(worksheet, Competitor.TableColumnNames(), row++);
                        foreach (var competitor in competitors)
                        {
                            PrintTableColumnData(worksheet, competitor.TableRowData(), row++);
                        }
                        row++;
                    }
                }
            }

            string firstCol = ExcelColumnFromNumber(1);
            string lastCol = ExcelColumnFromNumber(Competitor.TableColumnNames().Length);
            worksheet.Columns[firstCol + ":" + lastCol].AutoFit();
        }

        //protected void CreateFormsListWorksheet(NetOffice.ExcelApi.Worksheet worksheet)
        //{
        //    var row = 1;
        //    worksheet.Name = "Forms Divisions";
        //    PrintWorksheetTitle("Forms Divisions", worksheet, row++);

        //    var formsCompetitors = _reporter.GetCompetitors()
        //        .Where(x => x.IsForms);

        //    var competitorsByAgeDivision = EventInfo.AgeDivisions
        //            .ToDictionary(
        //                x => x,
        //                y => formsCompetitors.Where(z => z.IsIncludedInAgeDivision(y)));

        //    foreach (var ageGroup in competitorsByAgeDivision)
        //    {
        //        var competitorsByRank = ageGroup.Value
        //            .GroupBy(x => x.Rank);

        //        foreach (var rankGroup in competitorsByRank)
        //        {
        //            var competitorsByGender = rankGroup.GroupBy(x => x.Gender);
        //            foreach (var genderGroup in competitorsByGender)
        //            {
        //                var competitors = genderGroup.OrderBy(x => x.SchoolName);

        //                PrintGroupHeader(genderGroup.Key + ", " + ageGroup.Key.Name + ", " + rankGroup.Key, worksheet, row++);
        //                PrintTableColumnNames(worksheet, Competitor.TableColumnNames(), row++);
        //                foreach (var competitor in competitors)
        //                {
        //                    PrintTableColumnData(worksheet, competitor.TableRowData(), row++);
        //                }
        //                row++;
        //            }
        //        }
        //    }
        //    string firstCol = ExcelColumnFromNumber(1);
        //    string lastCol = ExcelColumnFromNumber(Competitor.TableColumnNames().Length);
        //    worksheet.Columns[firstCol + ":" + lastCol].AutoFit();
        //}


        protected void PrintWorksheetTitle(string title, NetOffice.ExcelApi.Worksheet worksheet, int row)
        {
            worksheet.Cells[row,"A"].Value = title;
            worksheet.Cells[row,"A"].Font.Size = HeaderFontSize;
            worksheet.Range(worksheet.Cells[row, "A"], worksheet.Cells[row, "F"]).Merge();
        }

        protected void PrintTableColumnNames(NetOffice.ExcelApi.Worksheet worksheet, string[] columnNames, int row)
        {
            int col = 1;
            foreach (var columnName in columnNames)
            {
                string excelCol = ExcelColumnFromNumber(col);
                worksheet.Cells[row, excelCol].Value = columnName;
                worksheet.Cells[row, excelCol].Font.Bold = true;
                col++;
            }
        }

        protected void PrintTableColumnData(NetOffice.ExcelApi.Worksheet worksheet, string[] columnData, int row)
        {
            int col = 1;
            foreach (var data in columnData)
            {
                string excelCol = ExcelColumnFromNumber(col);
                worksheet.Cells[row, excelCol].Value = data;
                col++;
            }
        }


        protected void PrintGroupHeader(string groupTitle, NetOffice.ExcelApi.Worksheet worksheet, int row)
        {
            worksheet.Cells[row, "A"].Value = groupTitle;
            worksheet.Cells[row, "A"].Font.Size = GroupFontSize;
            worksheet.Cells[row, "A"].Interior.Color = ColorToDouble(Colors.LightGray);
            worksheet.Range(worksheet.Cells[row, "A"], worksheet.Cells[row, "D"]).Merge();
        }

        protected string ExcelColumnFromNumber(int column)
        {
            string columnString = "";
            decimal columnNumber = column;
            while (columnNumber > 0)
            {
                decimal currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }
            return columnString;
        }

        protected double ColorToDouble(Color color)
        {
            uint returnValue = color.B;
            returnValue = returnValue << 8;
            returnValue += color.G;
            returnValue = returnValue << 8;
            returnValue += color.R;
            return returnValue;
        }
    }
}
