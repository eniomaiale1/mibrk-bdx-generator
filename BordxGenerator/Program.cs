using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Configuration;
using ClosedXML.Excel;
using BordxGenerator.Model;

namespace BordxGenerator
{
    class Program
    {
        static string fileTemplate = Path.GetFullPath(Properties.Settings.Default.BDXClaimTemplate);
        static string fileReports = Properties.Settings.Default.BDXClaimReports.EndsWith("\\") ? Properties.Settings.Default.BDXClaimReports : Properties.Settings.Default.BDXClaimReports + "\\";

        static bool PrepareFile(string fileName, DateTime dateProcess) {
            string workSheetName = dateProcess.ToString("yyyy-MM");
            string fileContex = fileReports + fileName + ".xlsx";
            XLWorkbook wbTemplate;
            IXLWorksheet wsTemplate;
            if (Properties.Settings.Default.BDXClaimTemplate != null && File.Exists(Properties.Settings.Default.BDXClaimTemplate))
            {
                wbTemplate = new XLWorkbook(fileTemplate); // load the existing excel file
                wsTemplate = wbTemplate.Worksheets.Worksheet("Template");
                
            }
            else {
                return false;
            }

            if (!File.Exists(fileContex)) {
                wsTemplate.Name = workSheetName;
                wbTemplate.SaveAs(fileReports + fileName + ".xlsx");
                return true;
            } else
            {
                var workbook = new XLWorkbook(fileContex);
                try
                {
                    workbook.Worksheets.Delete(workSheetName);
                }
                catch { }
                //workbook.Worksheets.Add("Template");
                wsTemplate.CopyTo(workbook, workSheetName);
                //workbook.Worksheets.Worksheet(workSheetName) = wsTemplate;
                workbook.Save();
                return true;
            }
        }

        static void ProcessData(List<ClaimBordx> data, string fileName, DateTime from, DateTime to, Period period) {
            int line = 10;

            var workbook = new XLWorkbook(fileReports + fileName + ".xlsx"); // load the existing excel file

            //workbook.Worksheets.Worksheet().Delete();

            var worksheet = workbook.Worksheets.Worksheet(from.ToString("yyyy-MM"));

            //Coverholder
            worksheet.Cell("B2").SetValue("Claims Bordereau");
            //Contract No (UMR).       
            worksheet.Cell("B3").SetValue(period.Contract);
            //Binder Period
            worksheet.Cell("B4").SetValue(period.From.ToShortDateString() + " - " + period.To.ToShortDateString());
            //Class of Business:
            worksheet.Cell("B5").SetValue("");
            //London Broker:  
            worksheet.Cell("B6").SetValue("Miller Insurance Services");
            //Reporting Month:       
            worksheet.Cell("B7").SetValue(from.ToString("MMMM"));

            foreach (ClaimBordx claim in data)
            {

                worksheet.Cell("A" + line).SetValue(claim.Insured);
                worksheet.Cell("B" + line).SetValue(claim.Address.Line1);
                worksheet.Cell("C" + line).SetValue(claim.Address.Zip);
                worksheet.Cell("D" + line).SetValue(claim.Claimant);
                worksheet.Cell("E" + line).SetValue(claim.Address.Line1);
                worksheet.Cell("F" + line).SetValue(claim.Address.Zip);
                worksheet.Cell("G" + line).SetValue(claim.PolicyNumber);
                worksheet.Cell("H" + line).SetValue(claim.Address.State);
                worksheet.Cell("I" + line).SetValue(claim.Address.Country);
                worksheet.Cell("J" + line).SetValue(claim.EffectiveDate == DateTime.MinValue ? "" : claim.EffectiveDate.ToShortDateString());
                worksheet.Cell("K" + line).SetValue(claim.ExpirationDate == DateTime.MinValue ? "" : claim.ExpirationDate.ToShortDateString());
                worksheet.Cell("L" + line).SetValue(claim.Year);
                worksheet.Cell("M" + line).SetValue(claim.LossDateFrom == DateTime.MinValue ? "" : claim.LossDateFrom.ToShortDateString());
                worksheet.Cell("N" + line).SetValue(claim.LossDateTo == DateTime.MinValue ? "" : claim.LossDateTo.ToShortDateString());
                worksheet.Cell("O" + line).SetValue(claim.ClaimNumber);
                worksheet.Cell("V" + line).SetValue(claim.LossDescription);
                worksheet.Cell("W" + line).SetValue(claim.LossLocation);
                worksheet.Cell("X" + line).SetValue(claim.OriginalCurrency);
                worksheet.Cell("Y" + line).SetValue(claim.SettlementCurrency);
                worksheet.Cell("Z" + line).SetValue(claim.AmountClaimed);
                worksheet.Cell("AA" + line).SetValue(claim.AmountPaid);
                worksheet.Cell("AB" + line).SetFormulaA1("=AA" + line);
                worksheet.Cell("AD" + line).SetFormulaA1("=AC" + line);
                worksheet.Cell("AE" + line).SetValue(claim.FeesPaid);
                worksheet.Cell("AF" + line).SetFormulaA1("=AE" + line);
                worksheet.Cell("AG" + line).SetFormulaA1("=+AA" + line + "-AC" + line + "+AE" + line);
                worksheet.Cell("AH" + line).SetFormulaA1("=+AB" + line + "-AD" + line + "+AF" + line);
                worksheet.Cell("AK" + line).SetFormulaA1("=+AI" + line + "+AJ" + line);
                worksheet.Cell("AL" + line).SetFormulaA1("=+AB" + line + "-AD" + line + "+AI" + line);
                worksheet.Cell("AM" + line).SetFormulaA1("=+AF" + line + "+AJ" + line);
                worksheet.Cell("AN" + line).SetFormulaA1("=+AL" + line + "+AM" + line);
                worksheet.Cell("AO" + line).SetValue(claim.DateClaimMade == DateTime.MinValue ? "" :claim.DateClaimMade.ToShortDateString());
                worksheet.Cell("AP" + line).SetValue(claim.DateClaimNotified == DateTime.MinValue ? "" : claim.DateClaimNotified.ToShortDateString());
                worksheet.Cell("AV" + line).SetValue(claim.DateClaimPaid == DateTime.MinValue ? "" : claim.DateClaimPaid.ToShortDateString());
                worksheet.Cell("AW" + line).SetValue(claim.DateFeesPaid == DateTime.MinValue ? "" : claim.DateFeesPaid.ToShortDateString());

                line++;

            }

            int line_end = line - 1;

            worksheet.Range("Z" + line, "AN" + line).Style.Fill.BackgroundColor = XLColor.Gray;
            worksheet.Range("Z" + line, "AN" + line).Style.Font.Bold = true;
            worksheet.Range("Z10", "AN" + line).Style.NumberFormat.Format = "#,##0.00";

            worksheet.Cell("Z" + line).SetValue("TOTALS");
            worksheet.Cell("AA" + line).SetFormulaA1("=SUM(AA10:AA" + line_end + ")");
            worksheet.Cell("AB" + line).SetFormulaA1("=SUM(AB10:AB" + line_end + ")");
            worksheet.Cell("AC" + line).SetFormulaA1("=SUM(AC10:AC" + line_end + ")");
            worksheet.Cell("AD" + line).SetFormulaA1("=SUM(AD10:AD" + line_end + ")");
            worksheet.Cell("AE" + line).SetFormulaA1("=SUM(AE10:AE" + line_end + ")");

            worksheet.Cell("AF" + line).SetFormulaA1("=SUM(AF10:AF" + line_end + ")");
            worksheet.Cell("AG" + line).SetFormulaA1("=SUM(AG10:AG" + line_end + ")");
            worksheet.Cell("AH" + line).SetFormulaA1("=SUM(AH10:AH" + line_end + ")");
            worksheet.Cell("AI" + line).SetFormulaA1("=SUM(AI10:AI" + line_end + ")");
            worksheet.Cell("AJ" + line).SetFormulaA1("=SUM(AJ10:AJ" + line_end + ")");
            worksheet.Cell("AK" + line).SetFormulaA1("=SUM(AK10:AK" + line_end + ")");
            worksheet.Cell("AL" + line).SetFormulaA1("=SUM(AL10:AL" + line_end + ")");
            worksheet.Cell("AM" + line).SetFormulaA1("=SUM(AM10:AM" + line_end + ")");
            worksheet.Cell("AN" + line).SetFormulaA1("=SUM(AN10:AN" + line_end + ")");

            //workbook.SaveAs(fileReports + DateTime.UtcNow.ToFileTime() + ".xlsx");
            workbook.Save();
        }

        static void Main(string[] args)
        {
            //Application excel = new Application();
            try
            {
                var today = DateTime.Today;
                var month = new DateTime(today.Year, today.Month, 1);
                var first = month.AddMonths(-1);
                var last = month.AddDays(-1);

                List<ClaimBordx> bordxData = DAL.GetReportData(first, last);
                List<Period> periods = DAL.GetPeriods();
                List<ClaimBordx> fees = DAL.GetFees(first, last);

                foreach (Period p in periods) {

                    List<ClaimBordx> periodData = bordxData.Where(c => (c.ExpirationDate <= p.To && c.ExpirationDate >= p.From)).ToList();
                    
                    periodData.AddRange(fees.Where(c => c.DateFeesPaid <= last && c.DateFeesPaid >= first && c.DateFeesPaid <= p.To && c.DateFeesPaid >= p.From));
                    string fileName = p.From.ToString("yyyyMMdd") + "-" + p.To.ToString("yyyyMMdd");

                    if (periodData.Count > 0 && PrepareFile(fileName, first))
                    {
                        ProcessData(periodData, fileName, first, last, p);
                    }
                }

            }
            catch (Exception es) { }
            finally{
                //excel.Quit();
            }
        }
    }
}
