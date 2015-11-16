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
                CreateFormsListWorksheet(workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[workbook.Worksheets.Count]) as NetOffice.ExcelApi.Worksheet);

                workbook.SaveAs(outFilePath);
                excelApp.Quit();
            }
        }

        protected void CreateCompetitorListWorksheet(NetOffice.ExcelApi.Worksheet worksheet)
        {
            var competitors = _reporter.GetCompetitors()
                .OrderBy(x => x.LastName);

            worksheet.Name = "Competitor List";
            PrintWorksheetTitle("Competitor List", worksheet);

            var row = 2;
            PrintTableColumnNames(worksheet, Competitor.TableFullColumnNames(), row++);
            foreach (var competitor in competitors)
            {
                PrintTableColumnData(worksheet, competitor.TableFullRowData(), row++);
            }
            string firstCol = ExcelColumnFromNumber(1);
            string lastCol = ExcelColumnFromNumber(Competitor.TableFullColumnNames().Length);
            worksheet.Columns[firstCol + ":" + lastCol].AutoFit();
        }

        protected void CreateSpectatorListWorksheet(NetOffice.ExcelApi.Worksheet worksheet)
        {
            var spectators = _reporter.GetSpectators()
                .OrderBy(x => x.LastName);

            worksheet.Name = "Spectator List";
            PrintWorksheetTitle("Spectator List", worksheet);

            var row = 2;
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

            worksheet.Name = "Referee List";
            PrintWorksheetTitle("Referee List", worksheet);

            var row = 2;
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

            worksheet.Name = "Coach List";
            PrintWorksheetTitle("Coach List", worksheet);

            var row = 2;
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
            var competitorsByRank = _reporter.GetCompetitors()
                .Where(x => x.IsSparring)
                .GroupBy(x => x.Rank);

            worksheet.Name = "Sparring Divisions";
            PrintWorksheetTitle("Sparring Divisions", worksheet);

            var row = 2;
            foreach (var rankGroup in competitorsByRank)
            {
                var competitorsByAgeDivision = rankGroup
                    .OrderBy(x => x.Age)
                    .GroupBy(x => x.AgeDivision);
                foreach (var ageGroup in competitorsByAgeDivision)
                {
                    var competitorsByGender = ageGroup.GroupBy(x => x.Gender);
                    foreach (var genderGroup in competitorsByGender)
                    {
                        var competitors = genderGroup.OrderBy(x => x.SchoolName);

                        PrintGroupHeader(genderGroup.Key + ", " + ageGroup.Key.Name + ", " + rankGroup.Key, worksheet, row++);
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

        protected void CreateFormsListWorksheet(NetOffice.ExcelApi.Worksheet worksheet)
        {
            var competitorsByRank = _reporter.GetCompetitors()
                .Where(x => x.IsForms)
                .GroupBy(x => x.Rank);

            worksheet.Name = "Forms Divisions";
            PrintWorksheetTitle("Forms Divisions", worksheet);

            var row = 2;
            foreach (var rankGroup in competitorsByRank)
            {
                var competitorsByAgeDivision = rankGroup
                    .OrderBy(x => x.Age)
                    .GroupBy(x => x.AgeDivision);
                foreach (var ageGroup in competitorsByAgeDivision)
                {
                    var competitorsByGender = ageGroup.GroupBy(x => x.Gender);
                    foreach (var genderGroup in competitorsByGender)
                    {
                        var competitors = genderGroup.OrderBy(x => x.SchoolName);

                        PrintGroupHeader(genderGroup.Key + ", " + ageGroup.Key.Name + ", " + rankGroup.Key, worksheet, row++);
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


        protected void PrintWorksheetTitle(string title, NetOffice.ExcelApi.Worksheet worksheet)
        {
            worksheet.Cells[1,"A"].Value = title;
            worksheet.Cells[1,"A"].Font.Size = HeaderFontSize;
            worksheet.Range(worksheet.Cells[1, "A"], worksheet.Cells[1, "F"]).Merge();
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
            worksheet.Range(worksheet.Cells[row, "A"], worksheet.Cells[row, "C"]).Merge();
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
