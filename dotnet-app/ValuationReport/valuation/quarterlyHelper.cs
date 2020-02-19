using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

class quarterlyHelper
{
    AsposHelper asposeHelper = new AsposHelper();
    static string staticholdingdescription;
    static string staticStartValue;
    public static bool ignoreReporting;
    public static Boolean isValSectionFound;
    public static int tieredFeeCount;
    public static Boolean tieredFeeScaleCheck;
    public static List<reconcileValuesProperty> invoiceReconcileList;
    public static List<reconcileValuesProperty> valuationReconcileList;
    public static List<reconcileValuesProperty> performanceReconcileList;
    public static double[] tieredFeeScaleList = { 10000000, 15000000, 25000000, 50000000, 400000000, 500000000 };


    public List<performanceproperty> getSymphonyPortfolioPerformance(string workBookPath, string ClientCode)
    {
        staticholdingdescription = null;
        staticStartValue = null;
        List<performanceproperty> oldreport = new List<performanceproperty>();
        int workSheetsCount = asposeHelper.getWorksheetsCount(workBookPath);
        bool doesPerformancePageExists = false;
        //var accountTypeShifted = false;
        
        for (int workSheetIterator = 0; workSheetIterator < workSheetsCount; workSheetIterator++)
        {

            doesPerformancePageExists = asposeHelper.searchStringWorksheet(workBookPath, 1, "Portfolio performance") && asposeHelper.searchStringWorksheet(workBookPath, 1, "Net Cash/Stock");
            var isContentsPage = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Contents");
            if (doesPerformancePageExists && !isContentsPage)
            {
                break;
            }

        }

        if (doesPerformancePageExists)
        {
            var tuplerowsColumn = asposeHelper.getRowsColumns(workBookPath, 1);
            var startRowPosistion = asposeHelper.getRowFromString(workBookPath, 1, "Portfolio Name");
            var portfolioColumnIndex = asposeHelper.getColumnIndexString(workBookPath, 1, "Portfolio Name");
            var startValueColumnIndex = asposeHelper.getColumnIndexString(workBookPath, 1, "Start Value");
            var netCashColumnIndex = asposeHelper.getColumnIndexString(workBookPath, 1, "Net Cash/Stock");
            var endValueColumnIndex = asposeHelper.getColumnIndexString(workBookPath, 1, "End Value");
            var appreciationColumnIndex = asposeHelper.getColumnIndexString(workBookPath, 1, "Appreciation/");
            var indexValuesColumnIndex = asposeHelper.getColumnIndexString(workBookPath, 1, "Index Values (Total Returns)");
            var getstartValueDateRow = asposeHelper.getRowFromString(workBookPath, 1, "Index Values (Total Returns)") + 2;

            var stopCollectingData = false;
            for (int cellIterator = startRowPosistion + 2; cellIterator < tuplerowsColumn.Item1; cellIterator++)
            {
                if (!stopCollectingData)
                {
                    var portfolioDescription = asposeHelper.getCellValue(workBookPath, 1, cellIterator, portfolioColumnIndex);
                    performanceproperty locallist = new performanceproperty();
                    if (!string.IsNullOrEmpty(portfolioDescription))
                    {
                        var startValue = asposeHelper.getCellValue(workBookPath, 1, cellIterator, startValueColumnIndex);
                        if (!string.IsNullOrEmpty(startValue))
                        {
                            staticholdingdescription = portfolioDescription;
                            staticStartValue = startValue;
                            var netCashvalue = asposeHelper.getCellValue(workBookPath, 1, cellIterator, netCashColumnIndex);
                            var endValue = asposeHelper.getCellValue(workBookPath, 1, cellIterator, endValueColumnIndex);
                            var appreciation = asposeHelper.getCellValue(workBookPath, 1, cellIterator, appreciationColumnIndex);
                            var threeMonths = asposeHelper.getCellValue(workBookPath, 1, cellIterator, appreciationColumnIndex + 1);
                            var tweleveMonths = asposeHelper.getCellValue(workBookPath, 1, cellIterator, appreciationColumnIndex + 2);
                            var startValueDate = asposeHelper.getCellValue(workBookPath, 1, getstartValueDateRow, indexValuesColumnIndex);
                            try
                            {
                                if (string.IsNullOrEmpty(startValueDate))
                                {
                                    startValueDate = asposeHelper.getCellValue(workBookPath, 1, startRowPosistion - 1, startValueColumnIndex - 2);
                                    if (string.IsNullOrEmpty(startValueDate))
                                    {
                                        startValueDate = asposeHelper.getCellValue(workBookPath, 1, startRowPosistion - 1, startValueColumnIndex - 1);
                                    }
                                    if (!string.IsNullOrEmpty(startValueDate))
                                    {
                                        var formattedStartDate = DateTime.ParseExact(startValueDate, "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM yy", CultureInfo.InvariantCulture);
                                        locallist.startValueDate = formattedStartDate;
                                    }
                                }
                                else if (startValueDate.IndexOf("/") > 0)
                                {
                                    var formattedStartDate = DateTime.ParseExact(startValueDate, "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM yy", CultureInfo.InvariantCulture);
                                    locallist.startValueDate = formattedStartDate;
                                }
                            }
                            catch
                            {

                            }
                            var endValueDate = asposeHelper.getCellValue(workBookPath, 1, startRowPosistion - 1, endValueColumnIndex);
                            try
                            {
                                var formattedEndDate = DateTime.ParseExact(endValueDate, "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM yy", CultureInfo.InvariantCulture);
                                locallist.endValueDate = formattedEndDate;
                            }
                            catch
                            {

                            }

                            locallist.portfolioName = portfolioDescription;
                            locallist.clientName = ClientCode;
                            if (!string.IsNullOrEmpty(startValue))
                            {
                                if (startValue.StartsWith("(") && startValue.EndsWith(")"))
                                {
                                    startValue = startValue.Replace("(", "-").Replace(")", "");
                                }
                                locallist.startValue = startValue;
                            }
                            if (!string.IsNullOrEmpty(netCashvalue))
                            {
                                if (netCashvalue.StartsWith("(") && netCashvalue.EndsWith(")"))
                                {
                                    netCashvalue = netCashvalue.Replace("(", "-").Replace(")", "");
                                }
                                locallist.netCash = netCashvalue;
                            }
                            if (!string.IsNullOrEmpty(endValue))
                            {
                                if (endValue.StartsWith("(") && endValue.EndsWith(")"))
                                {
                                    endValue = endValue.Replace("(", "-").Replace(")", "");
                                }

                                locallist.endValue = endValue;
                            }
                            if (!string.IsNullOrEmpty(appreciation))
                            {
                                if (appreciation.StartsWith("(") && appreciation.EndsWith(")"))
                                {
                                    appreciation = appreciation.Replace("(", "-").Replace(")", "");
                                }
                                locallist.appreciation = appreciation;
                            }

                            if (!portfolioDescription.ToLower().Contains("total"))
                            {
                                if (!string.IsNullOrEmpty(threeMonths))
                                {
                                    locallist.threeMonths = threeMonths;
                                }
                                if (!string.IsNullOrEmpty(tweleveMonths))
                                {
                                    locallist.twelveMonths = tweleveMonths;
                                }

                            }
                            oldreport.Add(locallist);
                        }
                        else
                        {
                            int index = oldreport.FindIndex(r => r.portfolioName == staticholdingdescription && r.startValue == staticStartValue);
                            if (index != -1)
                            {
                                if (Regex.IsMatch(portfolioDescription, @"^\d"))
                                {
                                    oldreport[index].portfolioName = staticholdingdescription + portfolioDescription;
                                }
                                else
                                {
                                    if (staticholdingdescription.Contains("-"))
                                    {
                                        oldreport[index].portfolioName = staticholdingdescription + portfolioDescription;
                                    }
                                    else
                                    {
                                        oldreport[index].portfolioName = staticholdingdescription + " " + portfolioDescription;
                                    }

                                }
                            }

                        }

                    }
                    if (portfolioDescription.Equals("Total", StringComparison.InvariantCultureIgnoreCase))
                    {
                        stopCollectingData = true;
                        break;
                    }
                }
            }
        }
        return oldreport;
    }
    public List<performanceproperty> getPortfolioPerformance(string workBookPath, string ClientCode, List<performanceproperty> symphonyPerformanceList)
    {
        staticholdingdescription = null;
        List<performanceproperty> newReport = new List<performanceproperty>();
        performanceReconcileList = new List<reconcileValuesProperty>();
        int workSheetsCount = asposeHelper.getWorksheetsCount(workBookPath);
        for (int workSheetIterator = 0; workSheetIterator < workSheetsCount; workSheetIterator++)
        {
            var doesPerformancePageExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Values and performance");
            var isContentsPage = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Contents");
            if (doesPerformancePageExists && !isContentsPage)
            {
                var tuplerowsColumn = asposeHelper.getRowsColumns(workBookPath, workSheetIterator);
                var portfoliostartRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Portfolio name");
                var portfolioColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Portfolio name");
                var netCashColumnIndex = portfolioColumnIndex + 2;
                var appreciationColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Increase /");
                var performanceColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Performance %");
                var performanceRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Individual portfolio performance");
                var threeMonthsColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "3 months");
                var tweleveMonthsColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "12 months");
                var stopCollectingData = false;
                List<string> localDb = new List<string>();
                for (int cellIterator = portfoliostartRowPosistion + 1; cellIterator < tuplerowsColumn.Item1; cellIterator++)
                {
                    if (!stopCollectingData)
                    {
                        performanceproperty locallist = new performanceproperty();
                        int startRow = 0;
                        int endRow = 0;
                        var portfolioDescription = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, portfolioColumnIndex);
                        //var startValueDate = asposeHelper.getCellValue(workBookPath, workSheetIterator, portfoliostartRowPosistion-1, portfolioColumnIndex);
                        bool portfolioPresentInSymphony = symphonyPerformanceList.Any(x => x.portfolioName == portfolioDescription);
                        var startValueDatePart1 = asposeHelper.getCellValue(workBookPath, workSheetIterator, portfoliostartRowPosistion - 1, portfolioColumnIndex);
                        var startValueDatePart2 = asposeHelper.getCellValue(workBookPath, workSheetIterator, portfoliostartRowPosistion - 1, portfolioColumnIndex + 1);
                        var startValueDate = startValueDatePart1 + "" + startValueDatePart2;
                        var endValueDate = asposeHelper.getCellValue(workBookPath, workSheetIterator, portfoliostartRowPosistion - 1, portfolioColumnIndex + 3);
                        locallist.startValueDate = startValueDate;
                        locallist.endValueDate = endValueDate;
                        bool portfolioPresentInlocalDB = localDb.Any(x => x != null && x == portfolioDescription);
                        if (portfolioPresentInSymphony && !portfolioPresentInlocalDB)
                        {
                            var startValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, portfolioColumnIndex + 1);
                            var netCashvalue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, netCashColumnIndex);
                            if (!string.IsNullOrEmpty(startValue) && !string.IsNullOrEmpty(netCashvalue))
                            {
                                staticholdingdescription = portfolioDescription;
                                locallist.portfolioName = portfolioDescription;
                                if (startValue.Contains("("))
                                {
                                    startValue = startValue.Replace("(", "-").Replace(")", "");

                                }
                                locallist.startValue = startValue;


                                netCashvalue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, netCashColumnIndex);
                                if (netCashvalue.Contains("("))
                                {
                                    netCashvalue = netCashvalue.Replace("(", "-").Replace(")", "");

                                }
                                locallist.netCash = netCashvalue;

                                locallist.clientName = ClientCode;
                                var endValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, netCashColumnIndex + 1);
                                if (endValue.Contains("("))
                                {
                                    endValue = endValue.Replace("(", "-").Replace(")", "");

                                }
                                locallist.endValue = endValue;


                                var appreciationvalue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, appreciationColumnIndex);
                                if (appreciationvalue.Contains("("))
                                {
                                    appreciationvalue = appreciationvalue.Replace("(", "-").Replace(")", "");

                                }
                                locallist.appreciation = appreciationvalue;

                                string threeMonthsvalue = null;
                                string twelveMonthsvalue = null;
                                locallist.threeMonths = threeMonthsvalue;
                                locallist.twelveMonths = twelveMonthsvalue;
                                newReport.Add(locallist);
                                var localReconcileList = new reconcileValuesProperty();
                                if (portfolioDescription == "Total")
                                {
                                    localReconcileList.sectionName = "performance";
                                    localReconcileList.totalMarketValue = Convert.ToDouble(endValue);
                                    performanceReconcileList.Add(localReconcileList);
                                }
                            }
                            localDb.Add(portfolioDescription);
                        }
                        else
                        {
                            startRow = cellIterator;
                            var nextDescription = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator + 1, portfolioColumnIndex);
                            if (string.IsNullOrEmpty(nextDescription))
                            {
                                nextDescription = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator + 2, portfolioColumnIndex);
                                endRow = cellIterator + 2;
                                if (portfolioDescription.Contains("-"))
                                {
                                    var d = portfolioDescription + nextDescription;
                                    locallist.portfolioName = d;
                                }
                                else
                                {
                                    var d = portfolioDescription + " " + nextDescription;
                                    locallist.portfolioName = d;
                                }

                            }
                            else
                            {
                                if (portfolioDescription.Contains("-"))
                                {
                                    var d = portfolioDescription + nextDescription;
                                    locallist.portfolioName = d;
                                }
                                else
                                {
                                    var d = portfolioDescription + " " + nextDescription;
                                    locallist.portfolioName = d;
                                }
                                endRow = cellIterator + 1;

                            }

                            localDb.Add(locallist.portfolioName);
                            // Now  get values from startRow and endRow
                            for (int j = startRow; j <= endRow; j++)
                            {
                                var startValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, j, portfolioColumnIndex + 1);
                                var netCashvalue = asposeHelper.getCellValue(workBookPath, workSheetIterator, j, netCashColumnIndex);
                                if (!string.IsNullOrEmpty(startValue) && !string.IsNullOrEmpty(netCashvalue))
                                {
                                    if (startValue.Contains("("))
                                    {
                                        startValue = startValue.Replace("(", "-").Replace(")", "");

                                    }
                                    locallist.startValue = startValue;
                                    //else
                                    //{
                                    //    locallist.startValue = startValue;
                                    //}

                                    netCashvalue = asposeHelper.getCellValue(workBookPath, workSheetIterator, j, netCashColumnIndex);
                                    if (netCashvalue.Contains("("))
                                    {
                                        netCashvalue = netCashvalue.Replace("(", "-").Replace(")", "");

                                    }
                                    locallist.netCash = netCashvalue;
                                    //else
                                    //{
                                    //    locallist.netCash = netCashvalue;
                                    //}

                                    locallist.clientName = ClientCode;
                                    var endValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, j, netCashColumnIndex + 1);
                                    if (endValue.Contains("("))
                                    {
                                        endValue = endValue.Replace("(", "-").Replace(")", "");

                                    }
                                    locallist.endValue = endValue;
                                    //else
                                    //{
                                    //    locallist.endValue = endValue;
                                    //}

                                    var appreciationvalue = asposeHelper.getCellValue(workBookPath, workSheetIterator, j, appreciationColumnIndex);
                                    if (appreciationvalue.Contains("("))
                                    {
                                        appreciationvalue = appreciationvalue.Replace("(", "-").Replace(")", "");

                                    }
                                    locallist.appreciation = appreciationvalue;
                                    //else
                                    //{
                                    //    locallist.appreciation = appreciationvalue;
                                    //}

                                }
                            }
                            cellIterator = endRow;
                            newReport.Add(locallist);
                        }
                        if (portfolioDescription.Equals("Total", StringComparison.InvariantCultureIgnoreCase))
                        {
                            stopCollectingData = true;
                            break;
                        }
                    }
                }

                break;
            }
        }
          return newReport;
    }
    public List<performanceproperty> comparePerformanceReport(List<performanceproperty> oldReport, List<performanceproperty> newReport)
    {
        staticholdingdescription = null;
        List<performanceproperty> missingList = new List<performanceproperty>();
        foreach (performanceproperty oldItem in oldReport)
        {
            var filteredList = newReport.Where(x => x.portfolioName.Equals(oldItem.portfolioName) && x.clientName != null && x.clientName.Equals(oldItem.clientName));
            performanceproperty locallist = new performanceproperty();
            if (filteredList.Count() > 0)
            {
                if (!string.IsNullOrEmpty(oldItem.portfolioName) && !string.IsNullOrEmpty(oldItem.netCash) && !string.IsNullOrEmpty(oldItem.startValue) && !string.IsNullOrEmpty(oldItem.appreciation) && !string.IsNullOrEmpty(oldItem.endValue) && !string.IsNullOrEmpty(filteredList.FirstOrDefault().portfolioName) && !string.IsNullOrEmpty(filteredList.FirstOrDefault().netCash) && !string.IsNullOrEmpty(filteredList.FirstOrDefault().startValue) && !string.IsNullOrEmpty(filteredList.FirstOrDefault().appreciation) && !string.IsNullOrEmpty(filteredList.FirstOrDefault().endValue))
                {
                    if (oldItem.portfolioName.Contains("Total"))
                    {
                        if (!oldItem.appreciation.Equals(filteredList.FirstOrDefault().appreciation, StringComparison.InvariantCultureIgnoreCase) || !oldItem.endValue.Equals(filteredList.FirstOrDefault().endValue, StringComparison.InvariantCultureIgnoreCase) || !oldItem.netCash.Equals(filteredList.FirstOrDefault().netCash, StringComparison.InvariantCultureIgnoreCase) || !oldItem.startValue.Equals(filteredList.FirstOrDefault().startValue, StringComparison.InvariantCultureIgnoreCase))
                        {
                            locallist.clientName = oldItem.clientName;
                            locallist.portfolioName = oldItem.portfolioName;
                            locallist.appreciation = oldItem.appreciation + "|" + filteredList.FirstOrDefault().appreciation;
                            locallist.endValue = oldItem.endValue + "|" + filteredList.FirstOrDefault().endValue;
                            locallist.netCash = oldItem.netCash + "|" + filteredList.FirstOrDefault().netCash;
                            locallist.startValue = oldItem.startValue + "|" + filteredList.FirstOrDefault().startValue;
                            locallist.threeMonths = oldItem.threeMonths + "|" + filteredList.FirstOrDefault().threeMonths;
                            locallist.twelveMonths = oldItem.twelveMonths + "|" + filteredList.FirstOrDefault().twelveMonths;
                            missingList.Add(locallist);
                        }

                    }
                    else if (!oldItem.appreciation.Equals(filteredList.FirstOrDefault().appreciation, StringComparison.InvariantCultureIgnoreCase) || !oldItem.endValue.Equals(filteredList.FirstOrDefault().endValue, StringComparison.InvariantCultureIgnoreCase) || !oldItem.netCash.Equals(filteredList.FirstOrDefault().netCash, StringComparison.InvariantCultureIgnoreCase) || !oldItem.startValue.Equals(filteredList.FirstOrDefault().startValue, StringComparison.InvariantCultureIgnoreCase) || oldItem.startValueDate != null && !oldItem.startValueDate.Equals(filteredList.FirstOrDefault().startValueDate, StringComparison.InvariantCultureIgnoreCase) || !oldItem.endValueDate.Equals(filteredList.FirstOrDefault().endValueDate, StringComparison.InvariantCultureIgnoreCase))
                    {
                        locallist.clientName = oldItem.clientName;
                        locallist.portfolioName = oldItem.portfolioName;
                        locallist.appreciation = oldItem.appreciation + "|" + filteredList.FirstOrDefault().appreciation;
                        locallist.endValue = oldItem.endValue + "|" + filteredList.FirstOrDefault().endValue;
                        locallist.netCash = oldItem.netCash + "|" + filteredList.FirstOrDefault().netCash;
                        locallist.startValue = oldItem.startValue + "|" + filteredList.FirstOrDefault().startValue;
                        locallist.startValueDate = oldItem.startValueDate + "|" + filteredList.FirstOrDefault().startValueDate;
                        locallist.endValueDate = oldItem.endValueDate + "|" + filteredList.FirstOrDefault().endValueDate;
                        missingList.Add(locallist);
                    }
                }
                else
                {
                    locallist.clientName = oldItem.clientName;
                    locallist.portfolioName = oldItem.portfolioName;
                    locallist.appreciation = "record missing in new generation report";
                    missingList.Add(locallist);
                }

            }

        }
        return missingList;
    }
    public List<invoiceproperty> getInvoiceDetails(string workBookPath, string ClientCode, string packType)
    {
        staticholdingdescription = null;

        List<invoiceproperty> newReport = new List<invoiceproperty>();      
        int workSheetsCount = asposeHelper.getWorksheetsCount(workBookPath);
        bool isTiered = false;
        tieredFeeCount = 0;
        List<string> tieredFeeList = new List<string>();
        string tieredValue = null;
        tieredFeeScaleCheck = true;
        invoiceReconcileList = new List<reconcileValuesProperty>();
        var multiPortfolio = false;
        double prevPortfolioValue = 0;

        for (int workSheetIterator = 0; workSheetIterator < workSheetsCount; workSheetIterator++)
        {
            
            var doesManagementFeePageExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Management fee invoice");
            var doesGlossaryExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Glossary");
            var isContentsPage = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Contents");
            if (doesManagementFeePageExists && !doesGlossaryExists && !isContentsPage)
            {
                var tuplerowsColumn = asposeHelper.getRowsColumns(workBookPath, workSheetIterator);
                for (int cellIterator = 1; cellIterator < tuplerowsColumn.Item1; cellIterator++)
                {
                    var feeDescriptionCoulmn = 2;
                    var valueCoulmn = feeDescriptionCoulmn + 1;
                    var feeValueColumn = 0;
                    var feeText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, feeDescriptionCoulmn);
                    if (!string.IsNullOrEmpty(feeText))
                    {
                        var locallist = new invoiceproperty();
                        if (feeText.Equals("Total"))
                         {

                            var valueText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, valueCoulmn);
                            if(valueText.StartsWith("("))
                            {
                                valueText = "-" + valueText.Replace("(", "").Replace(")", "").Trim();
                            }
                            var localReconcileList = new reconcileValuesProperty();
                            if (multiPortfolio)
                            {
                                int index = invoiceReconcileList.FindIndex(x => x.totalMarketValue == prevPortfolioValue);
                                var newTotalValue = Math.Round(prevPortfolioValue + Convert.ToDouble(valueText), 2);
                                invoiceReconcileList[index].totalMarketValue = newTotalValue;
                                prevPortfolioValue = newTotalValue;
                                //localReconcileList.totalMarketValue = Convert.ToDouble(valueText)
                            }
                            else
                            {
                                localReconcileList.sectionName = "invoice";
                                localReconcileList.totalMarketValue = Convert.ToDouble(valueText);
                                localReconcileList.valuesMatchCheck = false;
                                invoiceReconcileList.Add(localReconcileList);
                                prevPortfolioValue = localReconcileList.totalMarketValue;
                            }
                            

                            multiPortfolio = true;
                            
                        }
                        if (feeText.Equals("Zero fee", StringComparison.InvariantCultureIgnoreCase) || feeText.Equals("Ruffer Investment Management Fee", StringComparison.InvariantCultureIgnoreCase) || feeText.Equals("Ruffer Funds - No Additional Fee*", StringComparison.InvariantCultureIgnoreCase) || feeText.Equals("Ruffer Investment Company�", StringComparison.InvariantCultureIgnoreCase))
                        {
                            locallist.clientName = ClientCode;
                            locallist.feeDescription = feeText.Replace("�", "").Trim();
                            var valueText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, valueCoulmn);
                            if (!string.IsNullOrEmpty(valueText))
                            {
                                if (valueText.StartsWith("(") && valueText.EndsWith(")"))
                                {
                                    valueText = valueText.Replace("(", "-").Replace(")", "");
                                }
                                locallist.value = valueText.Replace("@", "").Trim();
                            }
                            var percentageText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, valueCoulmn + 1);
                            if (!string.IsNullOrEmpty(percentageText))
                            {
                                if (percentageText.ToLower().Contains("tiered"))
                                {
                                    isTiered = true;
                                    tieredValue = valueText;
                                }
                                if (percentageText.Contains("%") || percentageText.ToLower().Contains("tiered"))
                                {
                                    locallist.vat = percentageText.Replace("@", "").Trim();
                                    feeValueColumn = valueCoulmn + 2;
                                }
                                else
                                {
                                    percentageText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, valueCoulmn + 2);
                                    if (!string.IsNullOrEmpty(percentageText))
                                    {
                                        locallist.vat = percentageText.Replace("@", "").Trim();
                                        feeValueColumn = valueCoulmn + 3;
                                    }
                                }
                            }
                            var feeValueText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, feeValueColumn);
                            if (!string.IsNullOrEmpty(feeValueText))
                            {
                                locallist.fee = feeValueText.Trim();
                            }

                            newReport.Add(locallist);
                        }
                        else
                        {
                            feeText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, feeDescriptionCoulmn + 1);
                            if (!string.IsNullOrEmpty(feeText))
                            {
                                if (feeText.Equals("Zero fee", StringComparison.InvariantCultureIgnoreCase) || feeText.Equals("Ruffer Investment Management Fee", StringComparison.InvariantCultureIgnoreCase) || feeText.Equals("Ruffer Funds - No Additional Fee*", StringComparison.InvariantCultureIgnoreCase) || feeText.Equals("Ruffer Investment Company�", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    locallist.clientName = ClientCode;
                                    feeDescriptionCoulmn = 1 + 1;
                                    valueCoulmn = feeDescriptionCoulmn + 1;
                                    if (feeText.Equals("Normal Fees Apply", StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        locallist.feeDescription = "Ruffer Investment Management Fee";
                                    }
                                    else
                                    {
                                        locallist.feeDescription = feeText.Trim().Replace("�", "");
                                    }
                                    var valueText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, valueCoulmn);
                                    if (!string.IsNullOrEmpty(valueText))
                                    {
                                        if (valueText.StartsWith("(") && valueText.EndsWith(")"))
                                        {
                                            valueText = valueText.Replace("(", "-").Replace(")", "");
                                        }
                                        locallist.value = valueText.Trim().Replace("@", "");
                                    }
                                    var percentageText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, valueCoulmn + 1);
                                    if (!string.IsNullOrEmpty(percentageText))
                                    {
                                        if (percentageText.Contains("%"))
                                        {
                                            locallist.vat = percentageText.Trim().Replace("@", "");
                                            feeValueColumn = valueCoulmn + 2;
                                        }
                                        else
                                        {
                                            percentageText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, valueCoulmn + 2);
                                            if (!string.IsNullOrEmpty(valueText))
                                            {
                                                locallist.vat = percentageText.Trim().Replace("@", "");
                                                feeValueColumn = valueCoulmn + 3;
                                            }
                                        }
                                    }
                                    var feeValueText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, feeValueColumn);
                                    if (!string.IsNullOrEmpty(feeValueText))
                                    {
                                        locallist.fee = feeValueText.Trim();
                                    }
                                    newReport.Add(locallist);
                                }

                            }


                        }
                    }

                }
                var feeFromColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Fee from:");
                var feeFromRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Fee from:");
                if (feeFromColumnIndex > 0)
                {
                    var feeFromDescription = asposeHelper.getCellValue(workBookPath, workSheetIterator, feeFromRowPosistion, feeFromColumnIndex);
                    bool lessThanNintyDays = false;
                    if (feeFromDescription.Contains("(") && feeFromDescription.Contains("days"))
                    {

                        var tmp = feeFromDescription.Substring(feeFromDescription.IndexOf("(") + 1, (feeFromDescription.IndexOf(")") - 1) - feeFromDescription.IndexOf("(") + 1);

                        if (Int32.Parse(tmp.Split(" ").ToList()[0]) < 90)
                        {
                            lessThanNintyDays = true;
                        }

                    }
                    if (!string.IsNullOrEmpty(feeFromDescription) && !lessThanNintyDays)
                    {
                        var locallist = new invoiceproperty();
                        locallist.feeFrom = asposeHelper.getCellValue(workBookPath, workSheetIterator, feeFromRowPosistion, feeFromColumnIndex);
                        locallist.clientName = ClientCode;
                        newReport.Add(locallist);

                    }

                }
                var tieredFeeScaleExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Tiered fee scale");

                if (isTiered && tieredFeeScaleExists)
                {
                    var tieredFeeScale = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Tiered fee scale");
                    for (int cellIterator = tieredFeeScale; cellIterator < tuplerowsColumn.Item1; cellIterator++)
                    {
                        var tieredFeeScaleValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 1);
                        if (tieredFeeScaleValue.Contains("@"))
                        {
                            tieredFeeList.Add(tieredFeeScaleValue);
                        }
                    }
                }

            }

        }
        if (tieredFeeList.Count > 0)
        {
            int count = 0;
            if (!String.IsNullOrEmpty(tieredValue))
            {

                tieredValue = (tieredValue.Contains(",") ? tieredValue.Replace(",", "") : tieredValue);
                var tieredValueNumber = tieredValue.Contains(".") ? Double.Parse(tieredValue) : Convert.ToInt32(tieredValue);
                if (tieredValueNumber >= tieredFeeScaleList[0])
                {
                    count++;
                    var upperTierLimit = tieredFeeScaleList[0] + tieredFeeScaleList[1];
                    if (tieredValueNumber >= upperTierLimit)
                    {
                        count++;
                        upperTierLimit = upperTierLimit + tieredFeeScaleList[2];
                        if (tieredValueNumber >= upperTierLimit)
                        {
                            count++;
                            upperTierLimit = upperTierLimit + tieredFeeScaleList[3];
                            if (tieredValueNumber >= upperTierLimit)
                            {
                                count++;
                                upperTierLimit = upperTierLimit + tieredFeeScaleList[4];
                                if (tieredValueNumber >= upperTierLimit)
                                {
                                    count++;
                                    upperTierLimit = upperTierLimit + tieredFeeScaleList[5];
                                    if (tieredValueNumber >= upperTierLimit)
                                    {
                                        count++;
                                    }
                                }
                            }
                        }
                    }
                    count++;
                }
            }
            if (tieredFeeList.Count != count)
            {
                tieredFeeScaleCheck = false;
                tieredFeeCount = count;
            }

        }

        return newReport;
    }
    public List<invoiceproperty> getSymphonyInvoiceDetails(string workBookPath, string ClientCode, string packType)
    {
        staticholdingdescription = null;
        List<invoiceproperty> oldReport = new List<invoiceproperty>();
        int workSheetsCount = asposeHelper.getWorksheetsCount(workBookPath);
        for (int workSheetIterator = 0; workSheetIterator < workSheetsCount; workSheetIterator++)
        {
            var doesManagementFeePageExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Invoice");            
            var isContentsPage = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Contents");
            if (doesManagementFeePageExists && !isContentsPage)
            {
                var tuplerowsColumn = asposeHelper.getRowsColumns(workBookPath, workSheetIterator);
                for (int cellIterator = 1; cellIterator < tuplerowsColumn.Item1; cellIterator++)
                {
                    var feeDescriptionCoulmn = 1;
                    var valueCoulmn = feeDescriptionCoulmn + 1;
                    var feeValueColumn = 0;
                    var feeText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, feeDescriptionCoulmn);
                    if (!string.IsNullOrEmpty(feeText))
                    {
                        var locallist = new invoiceproperty();
                        if (feeText.Equals("Zero fee", StringComparison.InvariantCultureIgnoreCase) || feeText.Equals("Normal Fees Apply", StringComparison.InvariantCultureIgnoreCase) || feeText.Equals("Ruffer Funds - No Additional Fee*", StringComparison.InvariantCultureIgnoreCase) || feeText.Equals("Ruffer Investment Company", StringComparison.InvariantCultureIgnoreCase))
                        {
                            locallist.clientName = ClientCode;
                            if (feeText.Equals("Normal Fees Apply", StringComparison.InvariantCultureIgnoreCase) || feeText.Equals("Zero fee", StringComparison.InvariantCultureIgnoreCase))
                            {
                                locallist.feeDescription = "Ruffer Investment Management Fee";
                            }
                            else
                            {
                                locallist.feeDescription = feeText.Trim();
                            }
                            var valueText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, valueCoulmn);
                            if (!string.IsNullOrEmpty(valueText))
                            {
                                locallist.value = valueText.Trim().Replace("@", "");
                            }
                            var percentageText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, valueCoulmn + 1);
                            if (!string.IsNullOrEmpty(percentageText))
                            {
                                if (percentageText.Contains("%"))
                                {
                                    locallist.vat = percentageText.Trim().Replace("@", "");
                                    feeValueColumn = valueCoulmn + 2;
                                }
                                else
                                {
                                    percentageText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, valueCoulmn + 2);
                                    if (!string.IsNullOrEmpty(percentageText))
                                    {
                                        locallist.vat = percentageText.Trim().Replace("@", "");
                                        feeValueColumn = valueCoulmn + 3;
                                    }
                                }
                            }
                            var feeValueText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, feeValueColumn);
                            if (!string.IsNullOrEmpty(feeValueText))
                            {
                                locallist.fee = feeValueText.Trim();
                            }

                            oldReport.Add(locallist);
                        }
                        else
                        {
                            feeText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, feeDescriptionCoulmn + 1);
                            if (!string.IsNullOrEmpty(feeText))
                            {
                                if (feeText.Equals("Zero fee", StringComparison.InvariantCultureIgnoreCase) || feeText.Equals("Normal Fees Apply", StringComparison.InvariantCultureIgnoreCase) || feeText.Equals("Ruffer Funds - No Additional Fee*", StringComparison.InvariantCultureIgnoreCase) || feeText.Equals("Ruffer Investment Company", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    locallist.clientName = ClientCode;
                                    feeDescriptionCoulmn = 1 + 1;
                                    valueCoulmn = feeDescriptionCoulmn + 1;
                                    if (feeText.Equals("Normal Fees Apply", StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        locallist.feeDescription = "Ruffer Investment Management Fee";
                                    }
                                    else
                                    {
                                        locallist.feeDescription = feeText.Trim();
                                    }
                                    var valueText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, valueCoulmn);
                                    if (!string.IsNullOrEmpty(valueText))
                                    {
                                        locallist.value = valueText.Trim().Replace("@", "");
                                    }
                                    var percentageText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, valueCoulmn + 1);
                                    if (!string.IsNullOrEmpty(percentageText))
                                    {
                                        if (percentageText.Contains("%"))
                                        {
                                            locallist.vat = percentageText.Trim().Replace("@", "");
                                            feeValueColumn = valueCoulmn + 2;
                                        }
                                        else
                                        {
                                            percentageText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, valueCoulmn + 2);
                                            if (!string.IsNullOrEmpty(percentageText))
                                            {
                                                locallist.vat = percentageText.Trim().Replace("@", "");
                                                feeValueColumn = valueCoulmn + 3;
                                            }
                                        }
                                    }
                                    var feeValueText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, feeValueColumn);
                                    if (!string.IsNullOrEmpty(feeValueText))
                                    {
                                        locallist.fee = feeValueText.Trim();
                                    }

                                    oldReport.Add(locallist);
                                }

                            }


                        }
                    }
                }
                var feeFromColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Fee from:");
                var feeFromRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Fee from:");
                if (feeFromColumnIndex > 0)
                {
                    bool lessThanNintyDays = false;
                    var feeFromDescription = asposeHelper.getCellValue(workBookPath, workSheetIterator, feeFromRowPosistion, feeFromColumnIndex);
                    if (feeFromDescription.Contains("(") && feeFromDescription.Contains("days"))
                    {

                        var tmp = feeFromDescription.Substring(feeFromDescription.IndexOf("(") + 1, (feeFromDescription.IndexOf(")") - 1) - feeFromDescription.IndexOf("(") + 1);

                        if (Int32.Parse(tmp.Split(" ").ToList()[0]) < 90)
                        {
                            lessThanNintyDays = true;
                        }

                    }

                    if (!string.IsNullOrEmpty(feeFromDescription) && !lessThanNintyDays)
                    {
                        var fromDateStartIndex = feeFromDescription.LastIndexOf("from:") + 5;
                        var fromDateEndIndex = feeFromDescription.LastIndexOf(" to") - fromDateStartIndex;
                        var fromDate = feeFromDescription.Substring(fromDateStartIndex, fromDateEndIndex).Trim();
                        var formattedfromDate = DateTime.ParseExact(fromDate, "dd/MM/yy", CultureInfo.InvariantCulture).ToString("dd MMMM yyyy", CultureInfo.InvariantCulture);
                        var toDateStartIndex = feeFromDescription.LastIndexOf(" to") + 3;
                        var toDateEndIndex = feeFromDescription.LastIndexOf(" (") - toDateStartIndex;
                        var toDate = feeFromDescription.Substring(toDateStartIndex, toDateEndIndex).Trim();
                        var formattedtoDate = DateTime.ParseExact(toDate, "dd/MM/yy", CultureInfo.InvariantCulture).ToString("dd MMMM yyyy", CultureInfo.InvariantCulture);
                        var firstChange = feeFromDescription.Replace(fromDate, formattedfromDate.TrimStart('0'));
                        var secondChange = firstChange.Replace(toDate, formattedtoDate.TrimStart('0'));
                        var locallist = new invoiceproperty();
                        locallist.feeFrom = secondChange;
                        locallist.clientName = ClientCode;
                        oldReport.Add(locallist);

                    }
                }


            }

        }

        return oldReport;
    }
    public List<invoiceproperty> compareInvoice(List<invoiceproperty> oldReport, List<invoiceproperty> newReport)
    {
        staticholdingdescription = null;
        List<invoiceproperty> missingList = new List<invoiceproperty>();
        if (oldReport.Count > 0)
        {
            var oldReportWithoutFeeFrom = oldReport.Where(x => x.clientName != null && x.feeDescription != null && x.fee != null && x.value != null && x.vat != null);
            var newReportWithoutFeeFrom = newReport.Where(x => x.clientName != null && x.feeDescription != null && x.fee != null && x.value != null && x.vat != null);
            foreach (invoiceproperty oldItem in oldReportWithoutFeeFrom)
            {
                var dataExists = newReportWithoutFeeFrom.Any(x => x.clientName.Trim().Equals(oldItem.clientName.Trim(), StringComparison.InvariantCultureIgnoreCase) && x.feeDescription.Trim().Equals(oldItem.feeDescription.Trim(), StringComparison.InvariantCultureIgnoreCase) && x.fee.Trim().Equals(oldItem.fee.Trim(), StringComparison.InvariantCultureIgnoreCase) && x.value.Trim().Equals(oldItem.value.Trim(), StringComparison.InvariantCultureIgnoreCase) && x.vat.Trim().Equals(oldItem.vat.Trim(), StringComparison.InvariantCultureIgnoreCase));
                if (!dataExists)
                {
                    invoiceproperty locallist = new invoiceproperty();
                    locallist.clientName = oldItem.clientName;
                    locallist.feeDescription = oldItem.feeDescription;
                    locallist.fee = string.Format("Value : {0}, Percentage : {1} incorrect or missing in newGeneration report", oldItem.value, oldItem.vat);
                    missingList.Add(locallist);
                }
            }
            var oldReportOnlyFeeFrom = oldReport.Where(x => x.clientName != null && x.feeFrom != null);
            var newReportOnlyFeeFrom = newReport.Where(x => x.clientName != null && x.feeFrom != null);
            if (oldReportOnlyFeeFrom != null)
            {
                foreach (invoiceproperty oldFeeFromItem in oldReportOnlyFeeFrom)
                {
                    foreach (invoiceproperty newFeeFromItem in newReportOnlyFeeFrom)
                    {
                        if (oldFeeFromItem.clientName.Trim().Equals(newFeeFromItem.clientName.Trim(), StringComparison.InvariantCultureIgnoreCase))
                        {
                            if (!oldFeeFromItem.feeFrom.Equals(newFeeFromItem.feeFrom, StringComparison.InvariantCultureIgnoreCase))
                            {
                                invoiceproperty locallist = new invoiceproperty();
                                locallist.clientName = oldFeeFromItem.clientName;
                                locallist.feeFrom = oldFeeFromItem.feeFrom + "|" + newFeeFromItem.feeFrom;
                                missingList.Add(locallist);
                            }

                        }

                    }
                }

            }
        }

        return missingList;
    }
    public List<valuationproperty> getSymphonyQuarterlyValuationData(string workBookPath, string ClientCode)
    {
        // column index B-1 , C-2, D-3, E-4,F-5,G-6,H-7,I-8        
        List<valuationproperty> oldreport = new List<valuationproperty>();
        int workSheetsCount = asposeHelper.getWorksheetsCount(workBookPath);
        isValSectionFound = false;

        for (int workSheetIterator = 0; workSheetIterator < workSheetsCount; workSheetIterator++)
        {
            // check for valuation record in worksheet
            var doesValuationExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Valuation as at");            
            var isContentsPage = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Contents");
            if (doesValuationExists && !isContentsPage)
            {
                // get max rows from worksheet
                var tuplerowsColumn = asposeHelper.getRowsColumns(workBookPath, workSheetIterator);
                //  get starting row position
                var startRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Holding Description");
                var portfolioCodeRowIndex = asposeHelper.getRowFromString(workBookPath, workSheetIterator, ClientCode);
                var portfolioCodeColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, ClientCode);
                var portfolioCode = asposeHelper.getCellValue(workBookPath, workSheetIterator, portfolioCodeRowIndex, portfolioCodeColumnIndex);
                if (startRowPosistion == 0)
                {
                    startRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "HoldingDescription");
                }
                for (int cellIterator = startRowPosistion + 2; cellIterator < tuplerowsColumn.Item1; cellIterator++)
                {
                    var holdingdecription = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 1);
                    var isStringPresent = asposeHelper.searchStringWorksheet(@"testdata\assestdataset.xls", 0, holdingdecription);
                    if (!string.IsNullOrEmpty(holdingdecription) && !isStringPresent && holdingdecription.ToLower().Trim() != "total value of securities and cash" && !holdingdecription.ToLower().Trim().Contains("exchange rates used") && !holdingdecription.ToLower().Trim().Contains("page") && !holdingdecription.Equals("securities", StringComparison.InvariantCultureIgnoreCase))
                    {
                        var percentAgeofTotalValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 6);
                        var estimatedGrossIncome = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 8);

                        if (!string.IsNullOrEmpty(estimatedGrossIncome))
                        {
                            var holding = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 2);
                            if (!string.IsNullOrEmpty(holding) && holding.StartsWith("(") && holding.EndsWith(")"))
                            {
                                holding = holding.Replace("(", "-").Replace(")", "");
                            }
                            var bookCost = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 3);
                            var marketPrice = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 4);
                            var marketValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 5);

                            var grossYield = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 7);


                            var locallist = new valuationproperty();


                            if (!string.IsNullOrEmpty(holding))
                            {
                                holding = (decimal.Parse(holding)).ToString();
                            }

                            if (!String.IsNullOrEmpty(marketPrice))
                            {
                                var withOutCurrencyCode = Regex.Replace(marketPrice, "[^0-9.]", "");
                                var formattedPrice = decimal.Parse(withOutCurrencyCode);
                                marketPrice = Math.Round(formattedPrice, 2).ToString();

                            }

                            if (!String.IsNullOrEmpty(marketValue))
                            {
                                marketValue = decimal.Parse(marketValue).ToString();
                            }
                            if (!String.IsNullOrEmpty(bookCost))
                            {
                                if (bookCost.Contains(","))
                                {
                                    bookCost = bookCost.Replace(",", "");
                                }
                            }
                            if (!String.IsNullOrEmpty(grossYield))
                            {
                                locallist.grossyield = grossYield;
                            }
                            if (!String.IsNullOrEmpty(estimatedGrossIncome))
                            {
                                //// Give 2 digit tolerance 
                                //if (estimatedGrossIncome.Length > 3)
                                //{
                                //    estimatedGrossIncome = estimatedGrossIncome.Substring(0, estimatedGrossIncome.Length - 2);
                                //}
                                if (estimatedGrossIncome.Contains(","))
                                {
                                    estimatedGrossIncome = estimatedGrossIncome.Replace(",", "");
                                }
                            }

                            locallist.holdingdescription = holdingdecription;
                            locallist.holding = holding;
                            locallist.clientName = ClientCode;
                            locallist.bookcost = bookCost;
                            locallist.marketprice = marketPrice;
                            locallist.portfolioCode = portfolioCode;
                            locallist.marketvalue = marketValue;
                            locallist.estimatedgrossincome = estimatedGrossIncome;



                            oldreport.Add(locallist);
                        }
                    }
                }

            }

        }

        if (oldreport.Count == 0)
        {
            valuationproperty locallist = new valuationproperty();
            locallist.clientName = ClientCode;
            locallist.holdingdescription = "Symphony Format Issue";
            locallist.holding = "";
            oldreport.Add(locallist);
        }
        return oldreport;
    }
    public List<valuationproperty> getNewReportQuarterlyValuationData(string workBookPath, string ClientCode)
    {
        staticholdingdescription = null;
        valuationReconcileList = new List<reconcileValuesProperty>();
        // Holding Description - 1 , PortfolioCode - 2, Holding - 3, Market Price -4, Market Value - 5, Book Cost - 6, Percentage of total value -7, Gross yield -8, Estimated Gross Income -9
        List<valuationproperty> newReport = new List<valuationproperty>();
        int workSheetsCount = asposeHelper.getWorksheetsCount(workBookPath);
        //int workSheetsCount = 2;
        for (int workSheetIterator = 0; workSheetIterator < workSheetsCount; workSheetIterator++)
        {
            // check for valuation record
            var doesValuationExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Valuation");
            var doesPortfolioCodeExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Portfolio");
            var doesBookCostExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Book");
            var doesGlossaryExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Glossary");
            var isContentsPage = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Contents");
            if (doesValuationExists && doesPortfolioCodeExists && doesBookCostExists &&!isContentsPage && !doesGlossaryExists)
            {
                // get max rows from worksheet
                var tuplerowsColumn = asposeHelper.getRowsColumns(workBookPath, workSheetIterator);
                //  get starting row position
                var startRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Security description");
                for (int cellIterator = startRowPosistion + 1; cellIterator < tuplerowsColumn.Item1 - 1; cellIterator++)
                {
                    var holdingdecription = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 1);
                    // check for keyword in excel sheet which is in test data folder, if yes ignore that row in excel sheet
                    var isStringPresent = asposeHelper.searchStringWorksheet(@"testdata\assestdataset.xls", 0, holdingdecription);
                    if(holdingdecription.ToLower().Trim() == "total")
                    {
                        var marketValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 5);
                        var localReconcileList = new reconcileValuesProperty();
                        localReconcileList.sectionName = "valuation";
                        localReconcileList.totalMarketValue = Convert.ToDouble(marketValue);

                        valuationReconcileList.Add(localReconcileList);
                    }
                    if (!string.IsNullOrEmpty(holdingdecription) && holdingdecription.ToLower().Trim() != "total" && !isStringPresent && !holdingdecription.ToLower().Trim().Contains("days accrued interest of") && !holdingdecription.ToLower().Trim().Contains("exchange rates used") && holdingdecription.ToLower().Trim() != "exchange rates used:")
                    {
                        var estimatedGrossIncome = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 9);
                        var percentAgeofTotalValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 7);
                        if (!string.IsNullOrEmpty(estimatedGrossIncome))
                        {
                            var holding = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 3);
                            if (holding.StartsWith("(") && holding.EndsWith(")"))
                            {
                                holding = holding.Replace("(", "-").Replace(")", "");
                            }
                            holding = holding.Replace(",","");
                            staticholdingdescription = holdingdecription;
                            var bookCost = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 6);
                            var marketPrice = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 4);
                            var marketValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 5);
                            var grossYield = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 8);

                            var portfolioCodeColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Portfolio");
                            var portfolioCode = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, portfolioCodeColumnIndex);
                            var locallist = new valuationproperty();
                            locallist.holdingdescription = holdingdecription;
                            // Round off to 4 digits to match with symphont report
                            if (string.IsNullOrEmpty(holding))
                            {
                                locallist.holding = null;
                            }
                            else
                            {
                                if ((holding.Length - holding.IndexOf(".") - 1) > 4)
                                {
                                    locallist.holding = Math.Round(decimal.Parse(holding), 4).ToString();
                                }
                                else
                                {
                                    var r = (decimal.Parse(holding)).ToString();
                                    locallist.holding = r;
                                }
                            }

                            if (!String.IsNullOrEmpty(bookCost))
                            {
                                if (bookCost.Contains(","))
                                {
                                    bookCost = bookCost.Replace(",", "");
                                }
                                else if (bookCost.StartsWith("(") && bookCost.EndsWith(")"))
                                {
                                    bookCost = bookCost.Replace("(", "-").Replace(")", "");
                                }
                            }
                            if (!String.IsNullOrEmpty(marketPrice))
                            {
                                marketPrice = Regex.Replace(marketPrice, "[^0-9.]", "");
                            }
                            if (!String.IsNullOrEmpty(marketValue))
                            {
                                if (marketValue.StartsWith("(") && marketValue.EndsWith(")"))
                                {
                                    marketValue = marketValue.Replace("(", "-").Replace(")", "");
                                }
                                marketValue = decimal.Parse(marketValue).ToString();
                            }

                            if (!String.IsNullOrEmpty(grossYield))
                            {
                                locallist.grossyield = grossYield;

                            }
                            if (!String.IsNullOrEmpty(estimatedGrossIncome))
                            {
                                //// Give 2 digit tolerance 
                                //if (estimatedGrossIncome.Length > 3)
                                //{
                                //    estimatedGrossIncome = estimatedGrossIncome.Substring(0, estimatedGrossIncome.Length - 2);
                                //    estimatedGrossIncome = decimal.Parse(estimatedGrossIncome).ToString();
                                //}
                                if (estimatedGrossIncome.StartsWith("(") && estimatedGrossIncome.EndsWith(")"))
                                {
                                    estimatedGrossIncome = estimatedGrossIncome.Replace("(", "-").Replace(")", "");
                                }
                                estimatedGrossIncome = decimal.Parse(estimatedGrossIncome).ToString();
                            }

                            locallist.holdingdescription = holdingdecription;
                            locallist.holding = holding;
                            locallist.clientName = ClientCode;
                            locallist.bookcost = bookCost;
                            locallist.marketprice = marketPrice;
                            locallist.marketvalue = marketValue;
                            locallist.portfolioCode = portfolioCode;
                            locallist.estimatedgrossincome = estimatedGrossIncome;
                            newReport.Add(locallist);
                        }
                        else
                        {

                            if (!holdingdecription.ToLower().Trim().Contains("exchange rates"))
                            {
                                int index = newReport.FindIndex(r => r.holdingdescription == staticholdingdescription);
                                if (index != -1)
                                {
                                    if (Regex.IsMatch(holdingdecription, @"^\d"))
                                    {
                                        newReport[index].holdingdescription = staticholdingdescription + holdingdecription;
                                    }
                                    else
                                    {
                                        newReport[index].holdingdescription = staticholdingdescription + " " + holdingdecription;
                                    }
                                }

                            }
                        }

                    }
                }
            }

        }
        return newReport;
    }
    public List<valuationproperty> compareQuarterlyValuationReports(List<valuationproperty> oldReport, List<valuationproperty> newReport)
    {
        staticholdingdescription = null;
        List<valuationproperty> differenceList = new List<valuationproperty>();
        List<valuationproperty> missingList = new List<valuationproperty>();
        var oldWithoutNUllHolding = oldReport.Where(d => d.holding != null);
        var newWithoutNUllHolding = newReport.Where(d => d.holding != null);
        foreach (valuationproperty old in oldWithoutNUllHolding)
        {
            valuationproperty locallist = new valuationproperty();
            var recordExist = true;
            foreach (valuationproperty newItem in newWithoutNUllHolding)
            {
                if (old.holdingdescription.Equals(newItem.holdingdescription, StringComparison.InvariantCultureIgnoreCase) && old.holding.Equals(newItem.holding, StringComparison.InvariantCultureIgnoreCase) && old.portfolioCode.Equals(newItem.portfolioCode, StringComparison.InvariantCultureIgnoreCase))
                {
                    if (old.holding != newItem.holding || old.bookcost != newItem.bookcost || old.estimatedgrossincome != newItem.estimatedgrossincome || old.grossyield != newItem.grossyield || old.marketvalue != newItem.marketvalue || old.percentageoftotalvalue != newItem.percentageoftotalvalue)
                    {
                        locallist.holdingdescription = old.holdingdescription;
                        locallist.bookcost = old.bookcost + "|" + newItem.bookcost;
                        locallist.estimatedgrossincome = old.estimatedgrossincome + "|" + newItem.estimatedgrossincome;
                        locallist.grossyield = old.grossyield + "|" + newItem.grossyield;
                        locallist.holding = old.holding + "|" + newItem.holding;
                        locallist.marketvalue = old.marketvalue + "|" + newItem.marketvalue;
                        locallist.percentageoftotalvalue = old.percentageoftotalvalue + "|" + newItem.percentageoftotalvalue;
                        locallist.clientName = old.clientName;
                        differenceList.Add(locallist);
                    }
                    break;
                }
            }
            if (!recordExist)
            {
                missingList.Add(old);
            }
        }
        var oldWithNUllHolding = oldReport.Where(d => d.holding == null);
        var newWithNUllHolding = newReport.Where(d => d.holding == null);
        foreach (valuationproperty old in oldWithNUllHolding)
        {
            valuationproperty locallist = new valuationproperty();
            var recordExist = true;
            foreach (valuationproperty newItem in oldWithNUllHolding)
            {
                if (old.holdingdescription.Equals(newItem.holdingdescription, StringComparison.InvariantCultureIgnoreCase) && old.holding == null & (old.bookcost != null && old.bookcost.Equals(newItem.bookcost, StringComparison.InvariantCultureIgnoreCase)))
                {
                    if (old.estimatedgrossincome != newItem.estimatedgrossincome || old.grossyield != newItem.grossyield || old.marketvalue != newItem.marketvalue || old.percentageoftotalvalue != newItem.percentageoftotalvalue)
                    {
                        locallist.holdingdescription = old.holdingdescription;
                        locallist.bookcost = old.bookcost + "|" + newItem.bookcost;
                        locallist.estimatedgrossincome = old.estimatedgrossincome + "|" + newItem.estimatedgrossincome;
                        locallist.grossyield = old.grossyield + "|" + newItem.grossyield;
                        locallist.holding = old.holding + "|" + newItem.holding;
                        locallist.marketvalue = old.marketvalue + "|" + newItem.marketvalue;
                        locallist.percentageoftotalvalue = old.percentageoftotalvalue + "|" + newItem.percentageoftotalvalue;
                        locallist.clientName = old.clientName;
                        differenceList.Add(locallist);
                    }
                }
            }
            if (!recordExist)
            {
                missingList.Add(old);
            }
        }
        foreach (valuationproperty missingItem in missingList)
        {
            valuationproperty locallist = new valuationproperty();
            if (missingItem.holdingdescription.Equals("Symphony Format Issue", StringComparison.InvariantCultureIgnoreCase))
            {
                locallist.holdingdescription = "Symphony Format Issue";
                locallist.bookcost = "|dummy";
                differenceList.Add(locallist);
            }
            else
            {
                locallist.holdingdescription = missingItem.holdingdescription;
                locallist.clientName = missingItem.clientName;
                locallist.bookcost = "missing";
                differenceList.Add(locallist);
            }
        }
        return differenceList;
    }
    public List<acquisitionDisposals> getSymphonyQuarterlytAquisitionData(string workBookPath, string ClientCode)
    {
        staticholdingdescription = null;
        List<acquisitionDisposals> oldReport = new List<acquisitionDisposals>();
        // int workSheetsCount = asposeHelper.getWorksheetsCount(workBookPath);
        // for (int workSheetIterator = 0; workSheetIterator < workSheetsCount; workSheetIterator++)
        // {
        //     var doesAqusitionExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Stock Name");
        //     if (doesAqusitionExists)
        //     {
        //         // get max rows from worksheet
        //         var tuplerowsColumn = asposeHelper.getRowsColumns(workBookPath, workSheetIterator);
        //         //  get starting row position
        //         var startRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Stock Name");
        //         var dateColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Date & Time");
        //         var qunatitynColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Quantity");
        //         var priceColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Price");
        //         var fxColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "FX");
        //         for (int cellIterator = startRowPosistion + 2; cellIterator < tuplerowsColumn.Item1; cellIterator++)
        //         {
        //             var description = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, dateColumnIndex+2);
        //             if (!String.IsNullOrEmpty(description) && !description.Equals("PURCHASES", StringComparison.InvariantCultureIgnoreCase) && !description.Equals("SALES", StringComparison.InvariantCultureIgnoreCase) && !description.ToLower().Contains("all transactions are expressed in the account currency"))
        //             {
        //                 var date = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, dateColumnIndex);
        //                 if (!String.IsNullOrEmpty(date))
        //                 {
        //                     var locallist = new acquisitionDisposals();
        //                     locallist.securityDescription = description;
        //                     locallist.clientName = ClientCode;
        //                     staticholdingdescription = description;
        //                     var formattedDate = DateTime.ParseExact(date, "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM yy", CultureInfo.InvariantCulture);
        //                     locallist.date = formattedDate.TrimStart('0');
        //                     var time = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, dateColumnIndex+1);
        //                     var quantity = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, qunatitynColumnIndex);         
        //                     if (!String.IsNullOrEmpty(quantity))
        //                     {
        //                         if(Regex.Matches(quantity.Trim(),@"[a-zA-Z]").Count() >0)
        //                         {
        //                             var newQunatity = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, qunatitynColumnIndex+1);
        //                             decimal value;
        //                             if(Decimal.TryParse(newQunatity, out value))
        //                             {

        //                             }else{
        //                                 newQunatity = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, qunatitynColumnIndex+2);

        //                             }
        //                             if(!string.IsNullOrEmpty(newQunatity))
        //                             {
        //                                 locallist.quantity = ((long)decimal.Parse(newQunatity)).ToString();
        //                             }

        //                         }else{
        //                             locallist.quantity = ((long)decimal.Parse(quantity)).ToString();
        //                         }

        //                     }
        //                     var price = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, priceColumnIndex);
        //                     if (!String.IsNullOrEmpty(price))
        //                     {
        //                         if(Regex.Matches(price,@"[^0-9. , ]").Count()>0)
        //                         {
        //                             locallist.price = Regex.Replace(price, "[^0-9.]", "");
        //                         }else{
        //                             var newPrice = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, priceColumnIndex+1);
        //                             if(Regex.Matches(newPrice,@"[^0-9. , ]").Count()>0)
        //                             {
        //                                 locallist.price = Regex.Replace(newPrice, "[^0-9.]", "");
        //                             }else{
        //                                 newPrice = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, priceColumnIndex+2);
        //                                 locallist.price = Regex.Replace(newPrice, "[^0-9.]", "");
        //                             }

        //                         }
        //                     }
        //                     var fxrate = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 6);
        //                     if (!String.IsNullOrEmpty(fxrate))
        //                     {
        //                         // Agreed to not add fxrate to comparison
        //                         // locallist.fxrate = fxrate;
        //                     }
        //                     oldReport.Add(locallist);

        //                 }
        //                 else
        //                 {

        //                     int index = oldReport.FindIndex(r => r.securityDescription == staticholdingdescription);
        //                     if (index != -1)
        //                     {
        //                         if (Regex.IsMatch(description, @"^\d"))
        //                         {
        //                             if (char.IsLetter(oldReport[index].securityDescription[oldReport[index].securityDescription.Length - 1]))
        //                             {
        //                                 oldReport[index].securityDescription = staticholdingdescription + " " + description;
        //                             }
        //                             else
        //                             {

        //                                 oldReport[index].securityDescription = staticholdingdescription + description;
        //                             }
        //                         }
        //                         else
        //                         {
        //                             oldReport[index].securityDescription = staticholdingdescription + " " + description;
        //                         }
        //                     }

        //                 }

        //             }

        //         }


        //     }

        // }
        return oldReport;
    }
    public List<acquisitionDisposals> getNewReportQuarterlyAquisitionData(string workBookPath, string ClientCode)
    {
        staticholdingdescription = null;
        List<acquisitionDisposals> newReport = new List<acquisitionDisposals>();
        // int workSheetsCount = asposeHelper.getWorksheetsCount(workBookPath);
        // for (int workSheetIterator = 1; workSheetIterator < workSheetsCount; workSheetIterator++)
        // {
        //     var doesAqusitionExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "FX rate");
        //     if (doesAqusitionExists)
        //     {
        //         // get max rows from worksheet
        //         var tuplerowsColumn = asposeHelper.getRowsColumns(workBookPath, workSheetIterator);
        //         //  get starting row position
        //         var startRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Trade date");
        //         for (int cellIterator = startRowPosistion + 1; cellIterator < tuplerowsColumn.Item1 - 1; cellIterator++)
        //         {
        //             var description = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 2);
        //             if (!String.IsNullOrEmpty(description) && !description.ToLower().Contains("total consideration"))
        //             {
        //                 var date = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 1);
        //                 if (!String.IsNullOrEmpty(date))
        //                 {
        //                     staticholdingdescription = description;
        //                     var locallist = new acquisitionDisposals();
        //                     locallist.securityDescription = description;
        //                     locallist.clientName = ClientCode;
        //                     locallist.date = date;
        //                     if (!description.Equals("Acquisitions", StringComparison.InvariantCultureIgnoreCase) || !description.Equals("Disposals", StringComparison.InvariantCultureIgnoreCase) || description.ToLower().Contains("commission"))
        //                     {
        //                         var quantity = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 3);
        //                         if (!String.IsNullOrEmpty(quantity))
        //                         {
        //                             locallist.quantity = ((long)decimal.Parse(quantity)).ToString();
        //                         }
        //                         var price = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 4);
        //                         if (!String.IsNullOrEmpty(price))
        //                         {
        //                             locallist.price = Regex.Replace(price, "[^0-9.]", "");
        //                         }
        //                         var fxrate = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 5);
        //                         if (!String.IsNullOrEmpty(fxrate))
        //                         {
        //                             // Agreed to not add fxrate to comparison
        //                             //locallist.fxrate = fxrate;
        //                         }
        //                         newReport.Add(locallist);
        //                     }
        //                 }
        //                 else
        //                 {
        //                     int index = newReport.FindIndex(r => r.securityDescription == staticholdingdescription);
        //                     if (index != -1)
        //                     {
        //                         if (Regex.IsMatch(description, @"^\d"))
        //                         {
        //                             if (char.IsLetter(newReport[index].securityDescription[newReport[index].securityDescription.Length - 1]))
        //                             {
        //                                 newReport[index].securityDescription = staticholdingdescription + " " + description;
        //                             }
        //                             else
        //                             {
        //                                 newReport[index].securityDescription = staticholdingdescription + description;
        //                             }
        //                         }
        //                         else
        //                         {
        //                             newReport[index].securityDescription = staticholdingdescription + " " + description;
        //                         }
        //                     }
        //                 }


        //             }

        //         }


        //     }

        // }
        return newReport;
    }
    public List<acquisitionDisposals> compareQuarterlyAqusitionData(List<acquisitionDisposals> oldAquReport, List<acquisitionDisposals> newAquReport)
    {
        staticholdingdescription = null;
        List<acquisitionDisposals> missingList = new List<acquisitionDisposals>();
        foreach (acquisitionDisposals oldItem in oldAquReport)
        {
            acquisitionDisposals locallist = new acquisitionDisposals();
            var recordExists = newAquReport.Any(x => x.securityDescription.Equals(oldItem.securityDescription, StringComparison.InvariantCultureIgnoreCase) && x.clientName.Equals(oldItem.clientName, StringComparison.InvariantCultureIgnoreCase));
            if (recordExists)
            {
                var oldReportGroup = oldAquReport.Where(x => x.date != null && x.date.Equals(oldItem.date, StringComparison.InvariantCultureIgnoreCase) && (x.quantity != null && x.quantity.Equals(oldItem.quantity, StringComparison.InvariantCulture)));
                var newReportGroup = newAquReport.Where(x => x.date != null && x.date.Equals(oldItem.date, StringComparison.InvariantCultureIgnoreCase) && (x.quantity != null && x.quantity.Equals(oldItem.quantity, StringComparison.InvariantCulture)));
                if (oldReportGroup.Count() > 1)
                {
                    if (oldReportGroup.Count() != newReportGroup.Count())
                    {
                        locallist.clientName = oldItem.clientName;
                        locallist.securityDescription = oldItem.securityDescription;
                        locallist.date = newReportGroup.Count() + " record is present but it should be " + oldReportGroup.Count();
                        missingList.Add(locallist);
                    }
                }
                else if (oldReportGroup.Count() == 1 && newReportGroup.Count() == 1)
                {
                    if (!oldReportGroup.FirstOrDefault().date.IsEqualOrBothNullOrEmpty(newReportGroup.FirstOrDefault().date) && !oldReportGroup.FirstOrDefault().price.IsEqualOrBothNullOrEmpty(newReportGroup.FirstOrDefault().price) && !oldReportGroup.FirstOrDefault().quantity.IsEqualOrBothNullOrEmpty(newReportGroup.FirstOrDefault().quantity))
                    {
                        locallist.clientName = oldItem.clientName;
                        locallist.securityDescription = oldItem.securityDescription;
                        locallist.date = oldReportGroup.FirstOrDefault().date + "|" + newReportGroup.FirstOrDefault().date;
                        locallist.price = oldReportGroup.FirstOrDefault().price + "|" + newReportGroup.FirstOrDefault().price;
                        locallist.quantity = oldReportGroup.FirstOrDefault().quantity + "|" + newReportGroup.FirstOrDefault().quantity;
                        missingList.Add(locallist);
                    }
                }
                else
                {
                    locallist.securityDescription = oldItem.securityDescription;
                    locallist.date = "record missing in new generation report";
                    locallist.clientName = oldItem.clientName;
                    missingList.Add(locallist);
                }
            }
            else
            {
                locallist.securityDescription = oldItem.securityDescription;
                locallist.date = "record missing in new generation report";
                locallist.clientName = oldItem.clientName;
                missingList.Add(locallist);
            }
        }
        // // For aqusiyion alone before publishing missing list check with contains and then publish it
        // var toBeremoved  = missingList;
        // foreach (acquisitionDisposals missedItem in missingList)
        // {
        //     var oldItem = oldAquReport.Where(x => x.securityDescription.Equals(missedItem.securityDescription, StringComparison.InvariantCultureIgnoreCase) && x.clientName.Equals(missedItem.clientName, StringComparison.InvariantCultureIgnoreCase));
        //     var newItem = newAquReport.Where(x => x.securityDescription.Contains(missedItem.securityDescription) && x.clientName.Equals(missedItem.clientName, StringComparison.InvariantCultureIgnoreCase));
        //     if (newItem.Count() > 0)
        //     {
        //         toBeremoved.Remove(missedItem);

        //     }

        // }
        return missingList;
    }
    public List<cashproperty> getSymphonyQuarterlyCashStatements(string workBookPath, string ClientCode)
    {
        staticholdingdescription = null;
        // description - 1 , date - 2 , payments - 3 , receipts - 4, balance - 5
        List<cashproperty> oldCashReport = new List<cashproperty>();
        int workSheetsCount = asposeHelper.getWorksheetsCount(workBookPath);
        //var accountTypeShifted = false;
        for (int workSheetIterator = 0; workSheetIterator < workSheetsCount; workSheetIterator++)
        {
            // check for cash statements record in workSheets
            var cashSheetExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Payments");
            var balanceBroughtforwardExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Balance brought forward");
            var balanceCarriedForwardExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Balance carried forward");
            var dateColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Date");
            var descriptionColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Description");
            var paymentsColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Payments");
            var receiptsColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Receipts");
            var balanceColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Balance");
            var doesGlossaryExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Glossary");
            var isContentsPage = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Contents");
            if (cashSheetExists &&!isContentsPage &&!doesGlossaryExists)
            {
                if (balanceBroughtforwardExists)
                {
                    var startRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Balance brought forward");
                    var balanceColumn = asposeHelper.getCellValue(workBookPath, workSheetIterator, startRowPosistion, balanceColumnIndex);
                    var locallist = new cashproperty();
                    if (string.IsNullOrEmpty(balanceColumn))
                    {
                        var receiptsColumn = asposeHelper.getCellValue(workBookPath, workSheetIterator, startRowPosistion, receiptsColumnIndex);
                        if (string.IsNullOrEmpty(receiptsColumn))
                        {
                            var paymentCoulmn = asposeHelper.getCellValue(workBookPath, workSheetIterator, startRowPosistion, paymentsColumnIndex);
                            if (string.IsNullOrEmpty(paymentCoulmn))
                            {
                                var dateColumn = asposeHelper.getCellValue(workBookPath, workSheetIterator, startRowPosistion, dateColumnIndex);
                                if (!string.IsNullOrEmpty(dateColumn))
                                {
                                    locallist.description = "Balance brought forward";
                                    locallist.clientName = ClientCode;
                                    locallist.balance = dateColumn;
                                    oldCashReport.Add(locallist);
                                }
                            }
                            else
                            {
                                locallist.description = "Balance brought forward";
                                locallist.clientName = ClientCode;
                                locallist.balance = paymentCoulmn;
                                oldCashReport.Add(locallist);
                            }
                        }
                        else
                        {
                            locallist.description = "Balance brought forward";
                            locallist.clientName = ClientCode;
                            locallist.balance = receiptsColumn;
                            oldCashReport.Add(locallist);
                        }
                    }
                    else
                    {
                        locallist.description = "Balance brought forward";
                        locallist.clientName = ClientCode;
                        locallist.balance = balanceColumn;
                        oldCashReport.Add(locallist);
                    }
                }
                if (balanceCarriedForwardExists)
                {
                    var startRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Balance carried forward");
                    var balanceColumn = asposeHelper.getCellValue(workBookPath, workSheetIterator, startRowPosistion, balanceColumnIndex);
                    var locallist = new cashproperty();
                    if (string.IsNullOrEmpty(balanceColumn))
                    {
                        var receiptsColumn = asposeHelper.getCellValue(workBookPath, workSheetIterator, startRowPosistion, receiptsColumnIndex);
                        if (string.IsNullOrEmpty(receiptsColumn))
                        {
                            var paymentCoulmn = asposeHelper.getCellValue(workBookPath, workSheetIterator, startRowPosistion, paymentsColumnIndex);
                            if (string.IsNullOrEmpty(paymentCoulmn))
                            {
                                var dateColumn = asposeHelper.getCellValue(workBookPath, workSheetIterator, startRowPosistion, dateColumnIndex);
                                if (!string.IsNullOrEmpty(dateColumn))
                                {
                                    locallist.description = "Balance carried forward";
                                    locallist.clientName = ClientCode;
                                    locallist.balance = dateColumn;
                                    oldCashReport.Add(locallist);
                                }
                            }
                            else
                            {
                                locallist.description = "Balance carried forward";
                                locallist.clientName = ClientCode;
                                locallist.balance = paymentCoulmn;
                                oldCashReport.Add(locallist);
                            }
                        }
                        else
                        {
                            locallist.description = "Balance carried forward";
                            locallist.clientName = ClientCode;
                            locallist.balance = receiptsColumn;
                            oldCashReport.Add(locallist);
                        }
                    }
                    else
                    {
                        locallist.description = "Balance carried forward";
                        locallist.clientName = ClientCode;
                        locallist.balance = balanceColumn;
                        oldCashReport.Add(locallist);
                    }
                }

            }
        }
        return oldCashReport;
    }
    public List<cashproperty> getNewReportQuarterlyCashStatements(string workBookPath, string ClientCode)
    {
        staticholdingdescription = null;
        // Date - 1 , Description -2 , PaymentsGBP -3,ReceiptsGBP - 4, BalanceGBP - 5
        List<cashproperty> newCashReport = new List<cashproperty>();
        int workSheetsCount = asposeHelper.getWorksheetsCount(workBookPath);
        //Payments GBP
        for (int workSheetIterator = 0; workSheetIterator < workSheetsCount; workSheetIterator++)
        {
            // check for cash statements record in workSheets
            var cashSheetExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Payments GBP");
            var doesGlossaryExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Glossary");
            var isContentsPage = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Contents");
            if(!isContentsPage && !doesGlossaryExists)
            {
                if (!cashSheetExists)
                {
                    var descriptionRowIndex = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Description");
                    var cellvalue = asposeHelper.getCellValue(workBookPath, workSheetIterator, descriptionRowIndex, 3);
                    if (!string.IsNullOrEmpty(cellvalue) && cellvalue.ToLower().Trim().Contains("payments"))
                    {
                        cashSheetExists = true;
                    }
                }
                if (cashSheetExists)
                {
                    var tuplerowsColumn = asposeHelper.getRowsColumns(workBookPath, workSheetIterator);
                    var startRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Description");
                    for (int cellIterator = startRowPosistion + 1; cellIterator < tuplerowsColumn.Item1; cellIterator++)
                    {
                        var descriptionColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Description");
                        var cellValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, descriptionColumnIndex);
                        if (!string.IsNullOrEmpty(cellValue))
                        {

                            if (cellValue.Equals("Balance brought forward", StringComparison.InvariantCultureIgnoreCase))
                            {
                                var balanceColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Balance");
                                var balanceColumn = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, balanceColumnIndex);
                                var paymentsColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Payments");
                                var receiptsColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Receipts");
                                var dateColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Date");
                                var locallist = new cashproperty();
                                if (string.IsNullOrEmpty(balanceColumn))
                                {
                                    var receiptsColumn = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, receiptsColumnIndex);
                                    if (string.IsNullOrEmpty(receiptsColumn))
                                    {
                                        var paymentCoulmn = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, paymentsColumnIndex);
                                        if (string.IsNullOrEmpty(paymentCoulmn))
                                        {
                                            var dateColumn = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, dateColumnIndex);
                                            if (!string.IsNullOrEmpty(dateColumn))
                                            {
                                                locallist.description = "Balance brought forward";
                                                locallist.clientName = ClientCode;
                                                locallist.balance = dateColumn;
                                                newCashReport.Add(locallist);
                                            }
                                        }
                                        else
                                        {
                                            locallist.description = "Balance brought forward";
                                            locallist.clientName = ClientCode;
                                            locallist.balance = paymentCoulmn;
                                            newCashReport.Add(locallist);
                                        }
                                    }
                                    else
                                    {
                                        locallist.description = "Balance brought forward";
                                        locallist.clientName = ClientCode;
                                        locallist.balance = receiptsColumn;
                                        newCashReport.Add(locallist);
                                    }
                                }
                                else
                                {
                                    locallist.description = "Balance brought forward";
                                    locallist.clientName = ClientCode;
                                    locallist.balance = balanceColumn;
                                    newCashReport.Add(locallist);
                                }
                            }

                            if (cellValue.Equals("Balance carried forward", StringComparison.InvariantCultureIgnoreCase))
                            {
                                var balanceColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Balance");
                                var balanceColumn = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, balanceColumnIndex);
                                var paymentsColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Payments");
                                var receiptsColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Receipts");
                                var dateColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Date");

                                var locallist = new cashproperty();
                                if (string.IsNullOrEmpty(balanceColumn))
                                {
                                    var receiptsColumn = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, receiptsColumnIndex);
                                    if (string.IsNullOrEmpty(receiptsColumn))
                                    {
                                        var paymentCoulmn = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, paymentsColumnIndex);
                                        if (string.IsNullOrEmpty(paymentCoulmn))
                                        {
                                            var dateColumn = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, dateColumnIndex);
                                            if (!string.IsNullOrEmpty(dateColumn))
                                            {
                                                locallist.description = "Balance carried forward";
                                                locallist.clientName = ClientCode;
                                                locallist.balance = dateColumn;
                                                newCashReport.Add(locallist);
                                            }
                                        }
                                        else
                                        {
                                            locallist.description = "Balance carried forward";
                                            locallist.clientName = ClientCode;
                                            locallist.balance = paymentCoulmn;
                                            newCashReport.Add(locallist);
                                        }
                                    }
                                    else
                                    {
                                        locallist.description = "Balance carried forward";
                                        locallist.clientName = ClientCode;
                                        locallist.balance = receiptsColumn;
                                        newCashReport.Add(locallist);
                                    }
                                }
                                else
                                {
                                    locallist.description = "Balance carried forward";
                                    locallist.clientName = ClientCode;
                                    locallist.balance = balanceColumn;
                                    newCashReport.Add(locallist);
                                }


                            }

                        }
                    }
                }
            }
           


        }
        return newCashReport;
    }
    public List<cashproperty> compareQuarterlyCashStatements(List<cashproperty> oldCashReport, List<cashproperty> newCashReport)
    {
        staticholdingdescription = null;
        List<cashproperty> missingList = new List<cashproperty>();
        foreach (cashproperty oldItem in oldCashReport.Distinct())
        {
            cashproperty locallist = new cashproperty();
            var recordExists = newCashReport.Any(x => x.description.Equals(oldItem.description, StringComparison.InvariantCultureIgnoreCase) && x.clientName.Equals(oldItem.clientName, StringComparison.InvariantCultureIgnoreCase));
            if (recordExists)
            {
                if (oldItem.description.ToLower().Trim().Contains("balance"))
                {
                    var oldReportGroup = oldCashReport.Where(x => x.balance != null && x.balance.Equals(oldItem.balance, StringComparison.InvariantCultureIgnoreCase) && x.description.Equals(oldItem.description, StringComparison.InvariantCulture)).Select(f => f);
                    var newReportGroup = newCashReport.Where(x => x.balance != null && x.balance.Equals(oldItem.balance, StringComparison.InvariantCultureIgnoreCase) && x.description.Equals(oldItem.description, StringComparison.InvariantCulture)).Select(f => f);
                    if (oldReportGroup.Count() > 1)
                    {
                        if (oldReportGroup.Count() != newReportGroup.Count())
                        {
                            locallist.clientName = oldItem.clientName;
                            locallist.description = oldItem.description;
                            locallist.date = "only " + newReportGroup.Count() + " record is present but it should be " + oldReportGroup.Count();
                            missingList.Add(locallist);
                        }
                    }
                    else if (oldReportGroup.Count() == 1)
                    {
                        if (newReportGroup.Count() > 0)
                        {
                            if (!oldReportGroup.FirstOrDefault().balance.Equals(newReportGroup.FirstOrDefault().balance, StringComparison.InvariantCultureIgnoreCase))
                            {
                                locallist.clientName = oldItem.clientName;
                                locallist.description = oldItem.description;
                                locallist.balance = oldReportGroup.FirstOrDefault().balance + "|" + newReportGroup.FirstOrDefault().balance;
                                missingList.Add(locallist);
                            }
                        }
                        else
                        {
                            locallist.clientName = oldItem.clientName;
                            locallist.description = oldItem.description;
                            locallist.balance = oldItem.balance;
                            missingList.Add(locallist);
                        }
                    }
                    else
                    {
                        locallist.clientName = oldItem.clientName;
                        locallist.description = oldItem.description;
                        locallist.date = "check template in old report that might have affected comparison";
                        missingList.Add(locallist);
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(oldItem.date) && !string.IsNullOrEmpty(oldItem.paymentsReciepts))
                    {
                        var oldReportGroup = oldCashReport.Where(x => x.date != null && x.date.Equals(oldItem.date, StringComparison.InvariantCultureIgnoreCase) && x.description.Equals(oldItem.description, StringComparison.InvariantCulture) && x.paymentsReciepts.Equals(oldItem.paymentsReciepts, StringComparison.InvariantCulture));
                        var newReportGroup = newCashReport.Where(x => x.date != null && x.date.Equals(oldItem.date, StringComparison.InvariantCultureIgnoreCase) && x.description.Equals(oldItem.description, StringComparison.InvariantCulture) && x.paymentsReciepts.Equals(oldItem.paymentsReciepts, StringComparison.InvariantCulture));
                        if (oldReportGroup.Count() > 1)
                        {
                            if (oldReportGroup.Count() != newReportGroup.Count())
                            {
                                locallist.clientName = oldItem.clientName;
                                locallist.description = oldItem.description;
                                locallist.date = "only " + newReportGroup.Count() + " record is present but it should be " + oldReportGroup.Count();
                                missingList.Add(locallist);
                            }
                        }
                        else if (oldReportGroup.Count() == 1 && newReportGroup.Count() == 1)
                        {
                            if (!oldReportGroup.FirstOrDefault().date.IsEqualOrBothNullOrEmpty(newReportGroup.FirstOrDefault().date) && !oldReportGroup.FirstOrDefault().paymentsReciepts.IsEqualOrBothNullOrEmpty(newReportGroup.FirstOrDefault().paymentsReciepts))
                            {
                                locallist.clientName = oldItem.clientName;
                                locallist.description = oldItem.description;
                                locallist.date = oldReportGroup.FirstOrDefault().date + "|" + newReportGroup.FirstOrDefault().date;
                                locallist.paymentsReciepts = oldReportGroup.FirstOrDefault().paymentsReciepts + "|" + newReportGroup.FirstOrDefault().paymentsReciepts;
                                missingList.Add(locallist);
                            }
                        }
                        else
                        {
                            locallist.clientName = oldItem.clientName;
                            locallist.description = oldItem.description;
                            locallist.date = "record missing in new generation report dated - " + oldItem.date;
                            missingList.Add(locallist);
                        }
                    }
                    else
                    {
                        locallist.clientName = oldItem.clientName;
                        locallist.description = oldItem.description;
                        locallist.date = "check template in old report that might have affected comparison";
                        missingList.Add(locallist);
                    }
                }
            }
            else
            {
                if (oldItem.description.ToLower().Trim().Contains("gross interest calculated"))
                {
                    if (!string.IsNullOrEmpty(oldItem.paymentsReciepts) && ((long)decimal.Parse(oldItem.paymentsReciepts)) > 0)
                    {
                        locallist.clientName = oldItem.clientName;
                        locallist.description = oldItem.description;
                        locallist.date = "record missing in new generation report dated - " + oldItem.date;
                        missingList.Add(locallist);
                    }
                }
                else
                {
                    var finalCheckDescription = oldItem.description.Substring(0, (int)(oldItem.description.Length / 2));
                    var descriptionExists = newCashReport.Any(x => x.description.Contains(finalCheckDescription) && x.clientName.Equals(oldItem.clientName, StringComparison.InvariantCultureIgnoreCase));
                    if (!descriptionExists)
                    {
                        locallist.clientName = oldItem.clientName;
                        locallist.description = oldItem.description;
                        locallist.date = "record missing in new generation report dated - " + oldItem.date;
                        missingList.Add(locallist);
                    }
                }
            }
        }
        return missingList;
    }
    
    public List<reconcileValuesProperty> reconcileTotalValues()
    {
        List<reconcileValuesProperty> valuesList = new List<reconcileValuesProperty>();
        var localList = new reconcileValuesProperty();
        if (invoiceReconcileList.Count == 1 && valuationReconcileList.Count== 1)
        {
            if (invoiceReconcileList[0].totalMarketValue != valuationReconcileList[0].totalMarketValue)
            {
                if ((invoiceReconcileList[0].totalMarketValue - valuationReconcileList[0].totalMarketValue) > 0.1 || (invoiceReconcileList[0].totalMarketValue - valuationReconcileList[0].totalMarketValue) < -0.1)
                {
                    localList.valuesMatchCheck = true;
                    localList.sectionName = String.Format("Issue with market value on invoice page vs valuation : {0} | {1}", invoiceReconcileList[0].totalMarketValue, valuationReconcileList[0].totalMarketValue);
                    localList.totalMarketValue = invoiceReconcileList[0].totalMarketValue;
                    valuesList.Add(localList);
                }
            }
        }
        if (performanceReconcileList.Count == 1 && valuationReconcileList.Count == -1)
        { 
            if (performanceReconcileList[0].totalMarketValue != valuationReconcileList[0].totalMarketValue )
            {
                if ((performanceReconcileList[0].totalMarketValue - valuationReconcileList[0].totalMarketValue) > 0.1 || (performanceReconcileList[0].totalMarketValue - valuationReconcileList[0].totalMarketValue) < -0.1)
                {
                    localList = new reconcileValuesProperty();
                    localList.valuesMatchCheck = true;
                    localList.sectionName = String.Format("Issue with market value on performance page vs valuation: {0} | {1}", performanceReconcileList[0].totalMarketValue, valuationReconcileList[0].totalMarketValue);
                    localList.totalMarketValue = performanceReconcileList[0].totalMarketValue;
                    valuesList.Add(localList);
                }
            }
        }
        
        return valuesList;
    }
    public void createJsonQuarterlyReport(List<valuationproperty> valuationList, List<cashproperty> cashList, List<acquisitionDisposals> aqusitionList, List<performanceproperty> performanceList, List<invoiceproperty> invoiceList, List<reconcileValuesProperty> reconcileMarketValues, string outPutPath)
    {
        staticholdingdescription = null;
        List<quarterlyStruct> quarterlyList = new List<quarterlyStruct>();
        foreach (valuationproperty valuation in valuationList)
        {
            quarterlyStruct quarterly = new quarterlyStruct();
            quarterly.client = valuation.clientName;
            quarterly.valuationproperty = valuation;
            quarterlyList.Add(quarterly);
        }
        foreach (cashproperty cashItem in cashList)
        {
            quarterlyStruct quarterly = new quarterlyStruct();
            quarterly.client = cashItem.clientName;
            quarterly.cashproperty = cashItem;
            quarterlyList.Add(quarterly);
        }
        foreach (acquisitionDisposals acquistionItem in aqusitionList)
        {
            quarterlyStruct quarterly = new quarterlyStruct();
            quarterly.client = acquistionItem.clientName;
            quarterly.acquisitionDisposals = acquistionItem;
            quarterlyList.Add(quarterly);
        }
        foreach (performanceproperty performanceItem in performanceList)
        {

            quarterlyStruct quarterly = new quarterlyStruct();
            quarterly.client = performanceItem.clientName;
            quarterly.performance = performanceItem;
            quarterlyList.Add(quarterly);

        }
        foreach (invoiceproperty invoiceItem in invoiceList)
        {
            quarterlyStruct quarterly = new quarterlyStruct();
            quarterly.client = invoiceItem.clientName;
            quarterly.invoice = invoiceItem;
            quarterlyList.Add(quarterly);
        }
        foreach (reconcileValuesProperty reconcileItem in reconcileMarketValues)
        {
            quarterlyStruct quarterly = new quarterlyStruct();            
            quarterly.reconcileValues = reconcileItem;
            quarterlyList.Add(quarterly);
        }
        if (quarterlyList.Count > 0)
        {
            string jsonresponse = JsonConvert.SerializeObject(quarterlyList);
            System.IO.File.WriteAllText(outPutPath, jsonresponse);
        }

    }
    public void quarterlyReportTocsv(string folderPath, string outPutPath)
    {
        staticholdingdescription = null;
        Helper commonhelper = new Helper();
        List<string> checkZeroFee = new List<string>();
        var jsonFileList = commonhelper.getFileNames(folderPath, "*.json");
        var filePath = outPutPath;
        var csv = new StringBuilder();
        foreach (string symfile in jsonFileList)
        {
            List<quarterlyStruct> quarterlyList = JsonConvert.DeserializeObject<List<quarterlyStruct>>(File.ReadAllText(symfile));
            var fileName = Path.GetFileNameWithoutExtension(symfile); ;
            foreach (quarterlyStruct item in quarterlyList)
            {
                string bookCost = null;
                string clientName = null;
                string holdingDescription = null;
                string holding = null;
                // string marketprice = null;
                string marketValue = null;
                string grossyield = null;
                string estimatedgrossincome = null;
                string cashDescription = null;
                string cashDate = null;
                string paymentsreceipts = null;
                string cashBalance = null;
                string cashClientName = null;
                string aqusitionDescription = null;
                string aqusitionDate = null;
                string quantity = null;
                string price = null;
                string aqusitionClientName = null;
                string fxRate = null;
                string feeDescription = null;
                string feeFrom = null;
                string value = null;
                string fee = null;
                string feeTotal = null;
                string vat = null;
                string cummalitveTotal = null;
                string performanceclientName = null;
                string portfolioName = null; ;
                string startValue = null; ;
                string netCash = null; ;
                string endValue = null; ;
                string startValueDate = null;
                string endValueDate = null;
                string appreciation = null; ;
                string invoiceclientName = null;
                var valuationExist = false;
                var cashStatement = false;
                var aqusitionExist = false;
                var symphonyPDFbad = false;
                var performanceExist = false;
                var invoiceExist = false;
                var reconcileValuesIssue = false;
                var newReportAlignmnet = false;
                ignoreReporting = false;

                if (item.valuationproperty != null)
                {
                    bookCost = item.valuationproperty.bookcost;
                    clientName = item.valuationproperty.clientName;
                    holdingDescription = item.valuationproperty.holdingdescription;
                    holding = item.valuationproperty.holding;
                    marketValue = item.valuationproperty.marketvalue;
                    grossyield = item.valuationproperty.grossyield;
                    estimatedgrossincome = item.valuationproperty.estimatedgrossincome;
                    valuationExist = true;
                }
                if (item.cashproperty != null)
                {
                    cashDescription = item.cashproperty.description;
                    cashDate = item.cashproperty.date;
                    paymentsreceipts = item.cashproperty.paymentsReciepts;
                    cashBalance = item.cashproperty.balance;
                    cashClientName = item.cashproperty.clientName;
                    cashStatement = true;
                }
                if (item.acquisitionDisposals != null)
                {
                    aqusitionDescription = item.acquisitionDisposals.securityDescription;
                    aqusitionDate = item.acquisitionDisposals.date;
                    quantity = item.acquisitionDisposals.quantity;
                    price = item.acquisitionDisposals.price;
                    aqusitionClientName = item.acquisitionDisposals.clientName;
                    fxRate = item.acquisitionDisposals.fxrate;
                    aqusitionExist = true;
                }
                if (item.performance != null)
                {
                    portfolioName = item.performance.portfolioName;
                    startValue = item.performance.startValue;
                    netCash = item.performance.netCash;
                    endValue = item.performance.endValue;
                    appreciation = item.performance.appreciation;
                    performanceclientName = item.performance.clientName;
                    startValueDate = item.performance.startValueDate;
                    endValueDate = item.performance.endValueDate;
                    performanceExist = true;
                }
                if (item.invoice != null)
                {
                    feeDescription = item.invoice.feeDescription;
                    feeFrom = item.invoice.feeFrom;
                    value = item.invoice.value;
                    fee = item.invoice.fee;
                    feeTotal = item.invoice.feeTotal;
                    vat = item.invoice.vat;
                    cummalitveTotal = item.invoice.cummalitveTotal;
                    invoiceclientName = item.invoice.clientName;
                    invoiceExist = true;
                }
                if (item.reconcileValues != null)
                {
                    feeDescription = item.reconcileValues.sectionName;
                    value = item.reconcileValues.totalMarketValue.ToString();
                    reconcileValuesIssue = true;
                }

                if (valuationExist)
                {
                    Helper common = new Helper();
                    if (!string.IsNullOrEmpty(holding))
                    {
                        if (string.IsNullOrEmpty(common.splitString(holding, "|")[1]) && !string.IsNullOrEmpty(common.splitString(holding, "|")[0]))
                        {
                            newReportAlignmnet = true;
                        }
                    }
                    if (!string.IsNullOrEmpty(bookCost))
                    {
                        if (bookCost.Contains("|"))
                        {
                            if (string.IsNullOrEmpty(common.splitString(bookCost, "|")[1]) && !string.IsNullOrEmpty(common.splitString(bookCost, "|")[0]))
                            {
                                newReportAlignmnet = true;
                            }
                        }
                    }

                    if (!string.IsNullOrEmpty(marketValue))
                    {
                        if (string.IsNullOrEmpty(common.splitString(marketValue, "|")[1]) && !string.IsNullOrEmpty(common.splitString(marketValue, "|")[0]))
                        {
                            newReportAlignmnet = true;
                        }
                    }
                    if (!string.IsNullOrEmpty(grossyield))
                    {
                        if (string.IsNullOrEmpty(common.splitString(grossyield, "|")[1]) && !string.IsNullOrEmpty(common.splitString(grossyield, "|")[0]))
                        {
                            newReportAlignmnet = true;
                        }
                    }
                    if (!string.IsNullOrEmpty(estimatedgrossincome))
                    {
                        if (string.IsNullOrEmpty(common.splitString(estimatedgrossincome, "|")[1]) && !string.IsNullOrEmpty(common.splitString(estimatedgrossincome, "|")[0]))
                        {
                            newReportAlignmnet = true;
                        }
                    }
                }
                if (symphonyPDFbad)
                {
                    var section = "";
                    var description = "";
                    var issuedescription = "Symphony PDF Template Format Incorrect";
                    var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                    File.AppendAllText(filePath, newLine + Environment.NewLine);

                }
                if (newReportAlignmnet)
                {
                    var section = "";
                    var description = "";
                    var issuedescription = "Column Alignment Issue with new generation report";
                    var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                    File.AppendAllText(filePath, newLine + Environment.NewLine);
                }
                if (valuationExist && !symphonyPDFbad && !newReportAlignmnet)
                {
                    Helper common = new Helper();
                    if (!string.IsNullOrEmpty(holding))
                    {
                        if (holding.Contains("|"))
                        {
                            decimal oldReport;
                            decimal newReport;
                            Decimal.TryParse(common.splitString(holding, "|")[0], out oldReport);
                            Decimal.TryParse(common.splitString(holding, "|")[1], out newReport);
                            
                            if (oldReport != newReport)
                            {
                                var section = "valuation";
                                var description = "holding value";
                                var issuedescription = holding.Replace(",", "");
                                var deviation = "Difference   " + Math.Round(oldReport - newReport, 2).ToString();
                                var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription, deviation);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                            }
                        }
                        else
                        {
                            var section = "valuation";
                            var description = item.valuationproperty.holdingdescription.Replace(",", "");
                            var issuedescription = holding.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(bookCost))
                    {
                        ignoreReporting = false;
                        if (bookCost.Contains("|"))
                        {
                            double oldReport;
                            double newReport;
                            Double.TryParse(common.splitString(bookCost, "|")[0], out oldReport);
                            Double.TryParse(common.splitString(bookCost, "|")[1], out newReport);

                            var difference = Math.Round(oldReport - newReport, 2);

                            if (difference.ToString().Length > 1 && difference.GetType() == typeof(double) && !difference.ToString().Contains("-"))
                            {
                                ignoreReporting = difference <= 0.1 || difference <= -0.1;
                            }
                            if (oldReport != newReport && !ignoreReporting)
                            {
                                var section = "valuation";
                                var description = "bookCost value";
                                var issuedescription = bookCost.Replace(",", "");
                                var deviation = "Difference   " + difference.ToString();
                                var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription, deviation);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                            }
                        }
                        else
                        {
                            var section = "valuation";
                            var description = item.valuationproperty.bookcost.Replace(",", "");
                            var issuedescription = bookCost.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(marketValue))
                    {
                        ignoreReporting = false;
                        if (marketValue.Contains("|"))
                        {
                            double oldReport;
                            double newReport;
                            Double.TryParse(common.splitString(marketValue, "|")[0], out oldReport);
                            Double.TryParse(common.splitString(marketValue, "|")[1], out newReport);
                            var difference = Math.Round(oldReport - newReport, 2);

                            if (difference.ToString().Length > 1 && difference.GetType() == typeof(double))
                            {
                                ignoreReporting = difference <= 0.1 || difference <= -0.1;
                            }


                            if (oldReport != newReport && !ignoreReporting)
                            {
                                var section = "valuation";
                                var description = "Market value";
                                var issuedescription = marketValue.Replace(",", "") + " ( " + holdingDescription + " ) ";
                                var deviation = "Difference   " + difference.ToString();
                                var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription, deviation);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                            }
                        }
                        else
                        {
                            var section = "valuation";
                            var description = item.valuationproperty.marketvalue.Replace(",", "");
                            var issuedescription = marketValue.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(grossyield))
                    {
                        ignoreReporting = false;
                        if (grossyield.Contains("|"))
                        {
                            double oldReport;
                            double newReport;
                            Double.TryParse(common.splitString(grossyield, "|")[0], out oldReport);
                            Double.TryParse(common.splitString(grossyield, "|")[1], out newReport);

                            var difference = Math.Round(oldReport - newReport, 2);

                            if (difference.ToString().Length > 1 && difference.GetType() == typeof(double))
                            {
                                ignoreReporting = difference <= 0.1 || difference <= -0.1;
                            }

                            if (oldReport != newReport && !ignoreReporting)
                            {
                                var section = "valuation";
                                var description = "gross yield value";
                                var issuedescription = grossyield.Replace(",", "");
                                var deviation = "Difference   " + difference.ToString();
                                var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription, deviation);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                            }
                        }
                        else
                        {
                            var section = "valuation";
                            var description = item.valuationproperty.grossyield.Replace(",", "");
                            var issuedescription = grossyield.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(estimatedgrossincome))
                    {
                        if (estimatedgrossincome.Contains("|"))
                        {
                            double oldReport;
                            double newReport;
                            Double.TryParse(common.splitString(estimatedgrossincome, "|")[0], out oldReport);
                            Double.TryParse(common.splitString(estimatedgrossincome, "|")[1], out newReport);
                            var difference = Math.Round(oldReport - newReport, 2);

                            if (difference.ToString().Length >= 1 && difference.GetType() == typeof(double))
                            {

                                ignoreReporting = difference <= 1;
                            }
                            if (oldReport != newReport && !ignoreReporting)
                            {
                                var section = "valuation";
                                var description = "estimatedgrossincome value";
                                var issuedescription = estimatedgrossincome.Replace(",", "");
                                var deviation = "Difference   " + (oldReport - newReport).ToString();
                                var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription, deviation);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                            }
                        }
                        else
                        {
                            var section = "valuation";
                            var description = item.valuationproperty.estimatedgrossincome.Replace(",", "");
                            var issuedescription = estimatedgrossincome.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                }

                if (cashStatement && !symphonyPDFbad)
                {
                    Helper common = new Helper();
                    if (!string.IsNullOrEmpty(cashDate))
                    {
                        if (cashDate.Contains("|"))
                        {
                            if (!(common.splitString(cashDate, "|")[0]).Equals(common.splitString(cashDate, "|")[1], StringComparison.InvariantCultureIgnoreCase))
                            {
                                var section = "CashStatements";
                                var description = item.cashproperty.description.Replace(",", "");
                                var issuedescription = "cashDate value - " + common.splitString(cashDate, "|")[0].Replace(",", "") + "|" + common.splitString(cashDate, "|")[1].Replace(",", "");
                                var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                            }
                        }
                        else
                        {
                            var section = "CashStatements";
                            var description = item.cashproperty.description.Replace(",", "");
                            var issuedescription = cashDate.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(paymentsreceipts))
                    {
                        if (paymentsreceipts.Contains("|"))
                        {
                            if (!(common.splitString(paymentsreceipts, "|")[0]).Equals(common.splitString(paymentsreceipts, "|")[1], StringComparison.InvariantCultureIgnoreCase))
                            {
                                var section = "CashStatements";
                                var description = item.cashproperty.description.Replace(",", "");
                                var issuedescription = "paymentsreceipts value - " + common.splitString(paymentsreceipts, "|")[0].Replace(",", "") + "|" + common.splitString(paymentsreceipts, "|")[1].Replace(",", "");
                                var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                            }
                        }
                        else
                        {
                            var section = "CashStatements";
                            var description = item.cashproperty.description.Replace(",", "");
                            var issuedescription = paymentsreceipts.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(cashBalance))
                    {
                        if (cashBalance.Contains("|"))
                        {
                            if (!(common.splitString(cashBalance, "|")[0]).Equals(common.splitString(cashBalance, "|")[1], StringComparison.InvariantCultureIgnoreCase))
                            {
                                var section = "CashStatements";
                                var description = item.cashproperty.description.Replace(",", "");
                                var issuedescription = "cashBalance value - " + common.splitString(cashBalance, "|")[0].Replace(",", "") + "|" + common.splitString(cashBalance, "|")[1].Replace(",", "");
                                var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                            }
                        }
                        else
                        {
                            var section = "CashStatements";
                            var description = item.cashproperty.description.Replace(",", "");
                            var issuedescription = cashBalance.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                }
                if (aqusitionExist && !symphonyPDFbad)
                {
                    Helper common = new Helper();
                    if (!string.IsNullOrEmpty(aqusitionDate))
                    {
                        if (aqusitionDate.Contains("|"))
                        {
                            if (!(common.splitString(aqusitionDate, "|")[0]).Equals(common.splitString(aqusitionDate, "|")[1], StringComparison.InvariantCultureIgnoreCase))
                            {
                                var section = "Aqusition";
                                var description = item.acquisitionDisposals.securityDescription.Replace(",", "");
                                var issuedescription = "aqusitionDate value - " + common.splitString(aqusitionDate, "|")[0].Replace(",", "") + "|" + common.splitString(aqusitionDate, "|")[1].Replace(",", "");
                                var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                            }
                        }
                        else
                        {
                            var section = "Aqusition";
                            var description = item.acquisitionDisposals.securityDescription.Replace(",", "");
                            var issuedescription = aqusitionDate.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(quantity))
                    {
                        if (!(common.splitString(quantity, "|")[0]).Equals(common.splitString(quantity, "|")[1], StringComparison.InvariantCultureIgnoreCase))
                        {
                            var section = "Aqusition";
                            var description = item.acquisitionDisposals.securityDescription.Replace(",", "");
                            var issuedescription = "quantity value - " + common.splitString(quantity, "|")[0].Replace(",", "") + "|" + common.splitString(quantity, "|")[1].Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(price))
                    {
                        if (!(common.splitString(price, "|")[0]).Equals(common.splitString(price, "|")[1], StringComparison.InvariantCultureIgnoreCase))
                        {
                            var section = "Aqusition";
                            var description = item.acquisitionDisposals.securityDescription.Replace(",", "");
                            var issuedescription = "price value - " + common.splitString(price, "|")[0].Replace(",", "") + "|" + common.splitString(price, "|")[1].Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(fxRate))
                    {
                        if (!(common.splitString(fxRate, "|")[0]).Equals(common.splitString(fxRate, "|")[1], StringComparison.InvariantCultureIgnoreCase))
                        {
                            var section = "Aqusition";
                            var description = item.acquisitionDisposals.securityDescription.Replace(",", "");
                            var issuedescription = "fxRate value - " + common.splitString(fxRate, "|")[0].Replace(",", "") + "|" + common.splitString(fxRate, "|")[1].Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                }
                if (performanceExist && !symphonyPDFbad)
                {
                    Helper common = new Helper();
                    if (!string.IsNullOrEmpty(appreciation))
                    {
                        if (appreciation.Contains("|"))
                        {
                            double oldReport;
                            double newReport;
                            Double.TryParse(common.splitString(appreciation, "|")[0].Replace("(", "-").Replace(")", ""), out oldReport);
                            Double.TryParse(common.splitString(appreciation, "|")[1].Replace("(", "-").Replace(")", ""), out newReport);

                            var difference = Math.Round(oldReport - newReport, 2);

                            if (difference.ToString().Length > 1 && difference.GetType() == typeof(double))
                            {
                                ignoreReporting = difference <= 0.1 || difference <= -0.1;
                            }

                            if (!(common.splitString(appreciation, "|")[0]).Equals(common.splitString(appreciation, "|")[1], StringComparison.InvariantCultureIgnoreCase) && !ignoreReporting)
                            {
                                if (common.splitString(appreciation, "|")[1].Equals("�"))
                                {

                                }
                                else
                                {


                                    var section = "Performance";
                                    var description = item.performance.portfolioName.Replace(",", "");
                                    var issuedescription = "appreciation value - " + common.splitString(appreciation, "|")[0].Replace(",", "") + "|" + common.splitString(appreciation, "|")[1].Replace(",", "");
                                    var deviation = Math.Abs(float.Parse(common.splitString(appreciation, "|")[0].Replace("(", "").Replace(")", ""), CultureInfo.InvariantCulture.NumberFormat) - float.Parse(common.splitString(appreciation, "|")[1].Replace("(", "").Replace(")", ""), CultureInfo.InvariantCulture.NumberFormat));
                                    var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription, deviation);
                                    File.AppendAllText(filePath, newLine + Environment.NewLine);
                                }
                            }
                        }
                        else
                        {
                            var section = "Performance";
                            var description = item.performance.portfolioName.Replace(",", "");
                            var issuedescription = appreciation.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);

                        }
                    }
                    if (!string.IsNullOrEmpty(endValue))
                    {
                        if (endValue.Contains("|"))
                        {
                            double oldReport;
                            double newReport;
                            Double.TryParse(common.splitString(endValue, "|")[0].Replace("(", "-").Replace(")", ""), out oldReport);
                            Double.TryParse(common.splitString(endValue, "|")[1].Replace("(", "-").Replace(")", ""), out newReport);

                            var difference = Math.Round(oldReport - newReport, 2);

                            if (difference.ToString().Length > 1 && difference.GetType() == typeof(double))
                            {
                                ignoreReporting = difference <= 0.1 || difference <= -0.1;
                            }

                            if (!(common.splitString(endValue, "|")[0]).Equals(common.splitString(endValue, "|")[1], StringComparison.InvariantCultureIgnoreCase) && !ignoreReporting)
                            {
                                if (common.splitString(endValue, "|")[1].Equals("�"))
                                {

                                }
                                else
                                {

                                    var section = "Performance";
                                    var description = item.performance.portfolioName.Replace(",", "");
                                    var issuedescription = "endValue value - " + common.splitString(endValue, "|")[0].Replace(",", "") + "|" + common.splitString(endValue, "|")[1].Replace(",", "");
                                    var deviation = Math.Abs(float.Parse(common.splitString(endValue, "|")[0].Replace("(", "").Replace(")", ""), CultureInfo.InvariantCulture.NumberFormat) - float.Parse(common.splitString(endValue, "|")[1].Replace("(", "").Replace(")", ""), CultureInfo.InvariantCulture.NumberFormat));
                                    var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription, deviation);
                                    File.AppendAllText(filePath, newLine + Environment.NewLine);
                                }
                            }
                        }
                        else
                        {
                            var section = "Performance";
                            var description = item.performance.portfolioName.Replace(",", "");
                            var issuedescription = endValue.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(netCash))
                    {
                        if (netCash.Contains("|"))
                        {
                            if (!(common.splitString(netCash, "|")[0]).Equals(common.splitString(netCash, "|")[1], StringComparison.InvariantCultureIgnoreCase))
                            {
                                if (common.splitString(netCash, "|")[1].Equals("�"))
                                {

                                }
                                else
                                {

                                    var section = "Performance";
                                    var description = item.performance.portfolioName.Replace(",", "");
                                    var issuedescription = "netCash value - " + common.splitString(netCash, "|")[0].Replace(",", "") + "|" + common.splitString(netCash, "|")[1].Replace(",", "");
                                    var deviation = Math.Abs(float.Parse(common.splitString(netCash, "|")[0].Replace("(", "").Replace(")", ""), CultureInfo.InvariantCulture.NumberFormat) - float.Parse(common.splitString(netCash, "|")[1].Replace("(", "").Replace(")", ""), CultureInfo.InvariantCulture.NumberFormat));
                                    var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription, deviation);
                                    File.AppendAllText(filePath, newLine + Environment.NewLine);
                                }
                            }
                        }
                        else
                        {
                            var section = "Performance";
                            var description = item.performance.portfolioName.Replace(",", "");
                            var issuedescription = netCash.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(startValue))
                    {
                        if (startValue.Contains("|"))
                        {
                            if (!(common.splitString(startValue, "|")[0]).Equals(common.splitString(startValue, "|")[1], StringComparison.InvariantCultureIgnoreCase))
                            {
                                if (!common.splitString(startValue, "|")[1].Equals("�"))
                                {
                                    var section = "Performance";
                                    var description = item.performance.portfolioName.Replace(",", "");
                                    var issuedescription = "startValue value - " + common.splitString(startValue, "|")[0].Replace(",", "") + "|" + common.splitString(startValue, "|")[1].Replace(",", "");
                                    //var deviation = Math.Abs(float.Parse(common.splitString(startValue, "|")[0].Replace("(", "").Replace(")", ""), CultureInfo.InvariantCulture.NumberFormat) - float.Parse(common.splitString(startValue, "|")[1].Replace("(", "").Replace(")", ""), CultureInfo.InvariantCulture.NumberFormat));
                                    var deviation = Math.Round(float.Parse(common.splitString(startValue, "|")[0].Replace("(", "").Replace(")", ""), CultureInfo.InvariantCulture.NumberFormat) - float.Parse(common.splitString(startValue, "|")[1].Replace("(", "").Replace(")", ""), CultureInfo.InvariantCulture.NumberFormat),2);
                                    var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription, deviation);
                                    File.AppendAllText(filePath, newLine + Environment.NewLine);
                                }
                            }
                        }
                        else
                        {
                            var section = "Performance";
                            var description = item.performance.portfolioName.Replace(",", "");
                            var issuedescription = startValue.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }

                    if (!string.IsNullOrEmpty(startValueDate))
                    {
                        if (startValueDate.Contains("|"))
                        {
                            if (!(common.splitString(startValueDate, "|")[0]).Equals(common.splitString(startValueDate, "|")[1], StringComparison.InvariantCultureIgnoreCase))
                            {
                                bool notValidFormat = false;
                                DateTime dDate;
                                var date1 = common.splitString(startValueDate, "|")[0];
                                var date2 = common.splitString(startValueDate, "|")[1];
                                if (DateTime.TryParse(date1, out dDate))
                                {
                                    String.Format("{0:d/MM/yyyy}", dDate);
                                }
                                else
                                {
                                    notValidFormat = true;
                                }
                                if (DateTime.TryParse(date2, out dDate))
                                {
                                    String.Format("{0:d/MM/yyyy}", dDate);
                                }
                                else
                                {
                                    notValidFormat = true;
                                }

                                if (!notValidFormat)
                                {
                                    DateTime symphonyStartDate = DateTime.Parse((common.splitString(startValueDate, "|")[0]));
                                    DateTime newGenerationStartDate = DateTime.Parse((common.splitString(startValueDate, "|")[1]));
                                    var deviation = (symphonyStartDate - newGenerationStartDate).TotalDays;
                                    var section = "Performance";
                                    var description = item.performance.portfolioName.Replace(",", "");
                                    var issuedescription = "startValueDate value - " + common.splitString(startValueDate, "|")[0].Replace(",", "") + "|" + common.splitString(startValueDate, "|")[1].Replace(",", "");
                                    var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription, deviation);
                                    File.AppendAllText(filePath, newLine + Environment.NewLine);
                                }
                            }
                        }
                        else
                        {
                            var section = "Performance";
                            var description = item.performance.portfolioName.Replace(",", "");
                            var issuedescription = startValueDate.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(endValueDate))
                    {
                        if (endValueDate.Contains("|"))
                        {
                            if (!(common.splitString(endValueDate, "|")[0]).Equals(common.splitString(endValueDate, "|")[1], StringComparison.InvariantCultureIgnoreCase))
                            {
                                bool notValidFormat = false;
                                DateTime dDate;
                                var date1 = common.splitString(endValueDate, "|")[0];
                                var date2 = common.splitString(endValueDate, "|")[1];
                                if (DateTime.TryParse(date1, out dDate))
                                {
                                    String.Format("{0:d/MM/yyyy}", dDate);
                                }
                                else
                                {
                                    notValidFormat = true;
                                }
                                if (DateTime.TryParse(date2, out dDate))
                                {
                                    String.Format("{0:d/MM/yyyy}", dDate);
                                }
                                else
                                {
                                    notValidFormat = true;
                                }

                                if (!notValidFormat)
                                {
                                    DateTime symphonyDate = DateTime.Parse((common.splitString(endValueDate, "|")[0]));
                                    DateTime newGenerationDate = DateTime.Parse((common.splitString(endValueDate, "|")[1]));
                                    var deviation = (symphonyDate - newGenerationDate).TotalDays;
                                    var section = "Performance";
                                    var description = item.performance.portfolioName.Replace(",", "");
                                    var issuedescription = "endValueDate value - " + common.splitString(endValueDate, "|")[0].Replace(",", "") + "|" + common.splitString(endValueDate, "|")[1].Replace(",", "");
                                    var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription, deviation);
                                    File.AppendAllText(filePath, newLine + Environment.NewLine);
                                }
                            }
                        }
                        else
                        {
                            var section = "Performance";
                            var description = item.performance.portfolioName.Replace(",", "");
                            var issuedescription = endValueDate.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }

                }
                if (invoiceExist && !symphonyPDFbad)
                {
                    Helper common = new Helper();
                    if (!string.IsNullOrEmpty(fee))
                    {
                        var description = item.invoice.feeDescription.Replace(",", "");
                        if (description.Contains("Fee Exclusions -"))
                        {
                            checkZeroFee.Add(fileName);
                        }
                        else if (description.ToLower().Contains("totalwith fee "))
                        {
                            if (!checkZeroFee.Contains(fileName))
                            {
                                var section = "Invoice";
                                var issuedescription = fee.Replace(",", "");
                                var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                            }
                        }
                        else
                        {
                            var section = "Invoice";
                            var issuedescription = fee.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(cummalitveTotal))
                    {
                        if (cummalitveTotal.Contains("|"))
                        {
                            if (!(common.splitString(cummalitveTotal, "|")[0]).Equals(common.splitString(cummalitveTotal, "|")[1], StringComparison.InvariantCultureIgnoreCase))
                            {
                                var section = "Invoice";
                                var description = "Cummulative Total Incorrect";
                                var issuedescription = "cummulative value - " + common.splitString(cummalitveTotal, "|")[0].Replace(",", "") + "|" + common.splitString(cummalitveTotal, "|")[1].Replace(",", "");
                                var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                            }
                        }
                        else
                        {
                            var section = "Invoice";
                            var description = "Cummulative Total Incorrect";
                            var issuedescription = cummalitveTotal.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(feeFrom))
                    {
                        if (feeFrom.Contains("|"))
                        {
                            var symPhonyFreeFrom = common.splitString(feeFrom, "|")[0];
                            var newGenerationFreeFrom = common.splitString(feeFrom, "|")[1];
                            if (!string.IsNullOrEmpty(newGenerationFreeFrom) && !newGenerationFreeFrom.Equals("Ruffer", StringComparison.InvariantCultureIgnoreCase))
                            {
                                var symphonyDuration = symPhonyFreeFrom.Substring(symPhonyFreeFrom.LastIndexOf("(") + 1, (symPhonyFreeFrom.LastIndexOf("day") - 2) - symPhonyFreeFrom.LastIndexOf("(") + 1).Trim();
                                var newGenerationDuration = newGenerationFreeFrom.Substring(newGenerationFreeFrom.LastIndexOf("(") + 1, (newGenerationFreeFrom.LastIndexOf("day") - 2) - newGenerationFreeFrom.LastIndexOf("(") + 1).Trim();
                                if (Int32.Parse(symphonyDuration) == 0 && Int32.Parse(newGenerationDuration) > 0)
                                {
                                    var section = "Invoice";
                                    var description = "Zero Duration But Charged Wrongly";
                                    var issuedescription = "Fee From value - " + symPhonyFreeFrom.Replace(",", "") + "|" + newGenerationFreeFrom.Replace(",", "");
                                    var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                                    File.AppendAllText(filePath, newLine + Environment.NewLine);
                                }
                                else if (Int32.Parse(symphonyDuration) == Int32.Parse(newGenerationDuration))
                                {
                                    if (!symPhonyFreeFrom.Equals(newGenerationFreeFrom, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        var section = "Invoice";
                                        var description = "Total Days Same But Month Is Wrong";
                                        var issuedescription = "Fee From value - " + symPhonyFreeFrom.Replace(",", "") + "|" + newGenerationFreeFrom.Replace(",", "");
                                        var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                                        File.AppendAllText(filePath, newLine + Environment.NewLine);
                                    }
                                }
                                else
                                {
                                    if (!symPhonyFreeFrom.Equals(newGenerationFreeFrom, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        var section = "Invoice";
                                        var description = "Fee From incorrect";
                                        var issuedescription = "Fee From value - " + symPhonyFreeFrom.Replace(",", "") + "|" + newGenerationFreeFrom.Replace(",", "");
                                        var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                                        File.AppendAllText(filePath, newLine + Environment.NewLine);
                                    }
                                }
                            }
                            else
                            {
                                var section = "Invoice";
                                var description = "New Generation Fee Not Gathered , Check Your Code";
                                var issuedescription = feeFrom.Replace(",", "");
                                var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                            }
                        }
                        else
                        {
                            var section = "Invoice";
                            var description = "New Generation Fee Not Gathered , Check Your Code";
                            var issuedescription = feeFrom.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(feeTotal))
                    {
                        if (feeTotal.Contains("|"))
                        {
                            if (!(common.splitString(feeTotal, "|")[0]).Equals(common.splitString(feeTotal, "|")[1], StringComparison.InvariantCultureIgnoreCase))
                            {
                                var section = "Invoice";
                                var description = "Fee Total incorrect";
                                var issuedescription = "Fee Total value - " + common.splitString(feeTotal, "|")[0].Replace(",", "") + "|" + common.splitString(feeTotal, "|")[1].Replace(",", "");
                                var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                            }
                        }
                        else
                        {
                            var section = "Invoice";
                            var description = "Fee Total incorrect";
                            var issuedescription = feeTotal.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(vat))
                    {
                        if (vat.Contains("|"))
                        {
                            if (!(common.splitString(vat, "|")[0]).Equals(common.splitString(vat, "|")[1], StringComparison.InvariantCultureIgnoreCase))
                            {
                                var section = "Invoice";
                                var description = "VAT incorrect";
                                var issuedescription = "VAT value - " + common.splitString(vat, "|")[0].Replace(",", "") + "|" + common.splitString(vat, "|")[1].Replace(",", "");
                                var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                            }
                        }
                        else
                        {
                            var section = "Invoice";
                            var description = "VAT incorrect";
                            var issuedescription = vat.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!tieredFeeScaleCheck)
                    {
                        var section = "Invoice";
                        var issuedescription = "Tiered fee scale check";
                        var description = string.Format("Expected fee scale notes mismatch found: {0} instead", tieredFeeCount);
                        var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                        File.AppendAllText(filePath, newLine + Environment.NewLine);
                    }
                }
                if(reconcileValuesIssue)
                {
                    var diffInValues = feeDescription.Substring(feeDescription.IndexOf(":") + 1).Trim().Split("|");
                    double marketValue1;
                    double marketValue2;
                    Double.TryParse(diffInValues[0].Trim(), out marketValue1);
                    Double.TryParse(diffInValues[1].Trim(), out marketValue2);
                    var difference = Math.Round(marketValue1 - marketValue2, 2);
                    var deviation = "Difference in market value : " + difference.ToString();
                    var section = "Reconcile Market Values";
                    var issuedescription = feeDescription;
                    var description = "Total Market Value mismatch";
                    var newLine = string.Format("{0},{1},{2},{3}, {4}", fileName, section, description, issuedescription, deviation);
                    File.AppendAllText(filePath, newLine + Environment.NewLine);
                }
            }
        }
    }
    public void CompareQuarterlyReport(string symphonyPath, string newGenerationPath, string resultPath)
    {
        staticholdingdescription = null;
        Helper commonhelper = new Helper();
        AsposHelper asposeHelper = new AsposHelper();
        quarterlyHelper valuationHelper = new quarterlyHelper();
        MonthlyValuationHelper monthlyhelper = new MonthlyValuationHelper();
        var isFile = commonhelper.checkFileOrDirectory(symphonyPath);

        var missingRdbReportsFile = $@"{resultPath}\missingFiles.txt";

        if (isFile.Equals("file", StringComparison.InvariantCultureIgnoreCase))
        {
            var sympDirectoryPath = Path.GetDirectoryName(symphonyPath);
            var symFileName = Path.GetFileName(symphonyPath); ;
            var symFileLength = symFileName.Length;
            var lastIndex = symFileName.LastIndexOf("_");
            var clientName = symFileName.Substring(lastIndex + 1, symFileLength - (lastIndex + 1)).Replace(".pdf", "");
            var symFileNamewithoutExtension = Path.GetFileNameWithoutExtension(symFileName);
            var newGenFileName = Path.GetFileName(newGenerationPath);
            var newGenDirectoryPath = Path.GetDirectoryName(newGenerationPath);
            var newGenFileNamewithoutExtension = Path.GetFileNameWithoutExtension(newGenFileName);
            asposeHelper.convertPDFExcel(symphonyPath, sympDirectoryPath + symFileNamewithoutExtension + ".xls");
            asposeHelper.convertPDFExcel(newGenerationPath, newGenDirectoryPath + newGenFileNamewithoutExtension + ".xls");
            var oldValuationReport = valuationHelper.getSymphonyQuarterlyValuationData(sympDirectoryPath + symFileNamewithoutExtension + ".xls", clientName);
            var newValuationReport = valuationHelper.getNewReportQuarterlyValuationData(newGenDirectoryPath + newGenFileNamewithoutExtension + ".xls", clientName);
            var oldCashReport = valuationHelper.getSymphonyQuarterlyCashStatements(sympDirectoryPath + symFileNamewithoutExtension + ".xls", clientName);
            var newCashReport = valuationHelper.getNewReportQuarterlyCashStatements(newGenDirectoryPath + newGenFileNamewithoutExtension + ".xls", clientName);
            var oldAqusitionReport = valuationHelper.getSymphonyQuarterlytAquisitionData(sympDirectoryPath + symFileNamewithoutExtension + ".xls", clientName);
            var newAqusitionReport = valuationHelper.getNewReportQuarterlyAquisitionData(newGenDirectoryPath + newGenFileNamewithoutExtension + ".xls", clientName);
            var valuationResult = valuationHelper.compareQuarterlyValuationReports(oldValuationReport, newValuationReport);
            var cashResult = valuationHelper.compareQuarterlyCashStatements(oldCashReport, newCashReport);
            var aqusitionresult = valuationHelper.compareQuarterlyAqusitionData(oldAqusitionReport, newAqusitionReport);
            var oldPerformceResult = valuationHelper.getSymphonyPortfolioPerformance(sympDirectoryPath + symFileNamewithoutExtension + ".xls", clientName);
            var newPerformceResult = valuationHelper.getPortfolioPerformance(newGenDirectoryPath + newGenFileNamewithoutExtension + ".xls", clientName, oldPerformceResult);
            var performceResult = valuationHelper.comparePerformanceReport(oldPerformceResult, newPerformceResult);
            // var oldInvoiceResult = valuationHelper.getSymphonyInvoiceDetails(sympDirectoryPath + symFileNamewithoutExtension + ".xls", clientName, packtype);
            // var newInvoiceResult = valuationHelper.getInvoiceDetails(newGenDirectoryPath + newGenFileNamewithoutExtension + ".xls", clientName , packtype);
            // var invoiceResult = valuationHelper.compareInvoice(oldInvoiceResult,newInvoiceResult);
            // valuationHelper.createJsonQuarterlyReport(valuationResult, cashResult, aqusitionresult,performceResult,invoiceResult,Path.Combine(resultPath, symFileNamewithoutExtension) + ".json");

        }
        else if (isFile.Equals("directory", StringComparison.InvariantCultureIgnoreCase))
        {
            // First get all 
            var symphonyFileList1 = commonhelper.getFileNames(symphonyPath, "*.pdf");

            Helper helper = new Helper();
            // var symphonyFileList = symphonyFileList1.Where(a => !a.Contains("INSTSHADOW"));
            var symphonyFileList = symphonyFileList1;
            var newgenerationFileList = commonhelper.getFileNames(newGenerationPath, "*.pdf");
            List<string> client = new List<string>();
            int h = 0;
            // for debug to be removed
            var symphFileCount = symphonyFileList.Count();
            Dictionary<string, bool> rdbFileFoundLookup = new Dictionary<string, bool>();

            foreach (string symfile in symphonyFileList)
            {
                h = h + 1;
                Console.WriteLine("Total Files To Compare - " + symphFileCount);
                Console.WriteLine("Remaining file to process - " + (symphFileCount - h));
                string symFileName = symfile;
                string clientName = null;
                string packtype = null;
                if (symFileName.Contains("INSTSHADOW"))
                {
                    packtype = "INSTSHADOW";
                }
                else if (symFileName.Contains("LLQVALON"))
                {
                    packtype = "LLQVALON";
                }
                else if (symFileName.Contains("Q VAL NMA"))
                {
                    packtype = "Q VAL NMA";
                }
                else if (symFileName.Contains("Q VAL OFF"))
                {
                    packtype = "Q VAL OFF";
                }
                else if (symFileName.Contains("Q VAL SHAD"))
                {
                    packtype = "Q VAL SHAD";
                }
                else if (symFileName.Contains("QVALNMAOFF"))
                {
                    packtype = "QVALNMAOFF";
                }
                else if (symFileName.Contains("QVALOFFTAX"))
                {
                    packtype = "QVALOFFTAX";
                }
                else if (symFileName.Contains("QVALON"))
                {
                    packtype = "QVALON";
                }
                else if (symFileName.Contains("QVALONCON"))
                {
                    packtype = "QVALONCON";
                }
                else if (symFileName.Contains("QVALONINT"))
                {
                    packtype = "QVALONINT";
                }
                else if (symFileName.Contains("QVALPROF"))
                {
                    packtype = "QVALPROF";
                }
                else if (symFileName.Contains("M VAL BD2"))
                {
                    packtype = "M VAL BD2";
                }
                //if (symFileName.Contains("_ADD"))
                //{
                //    var lastIndexofADD = symFileName.LastIndexOf("_ADD");
                //    var newName = symFileName.Substring(0, lastIndexofADD);
                //    var lastIndex = newName.LastIndexOf("_");
                //    clientName = newName.Substring(lastIndex + 1, newName.Length - (lastIndex + 1)).Replace(".pdf", "");
                //}
                //else
                //{
                //    var lastIndex = symFileName.LastIndexOf("_");
                //    var firstIndex = symFileName.IndexOf("_");
                //    var symFileLength = symFileName.Length;
                //    clientName = symFileName.Substring(lastIndex + 1, symFileLength - (lastIndex + 1)).Replace(".pdf", "");
                //    //clientName = symFileName.Substring(firstIndex + 1, lastIndex - 1 - firstIndex).Replace(".pdf", "");
                //}

                var lastIndex = symFileName.LastIndexOf("_");
                var firstIndex = symFileName.IndexOf("_");
                var symFileLength = symFileName.Length;
                clientName = symFileName.Substring(lastIndex + 1, symFileLength - (lastIndex + 1)).Replace(".pdf", "");
                Directory.CreateDirectory(Path.Combine(resultPath, "jsonResults"));
                if (!client.Contains(clientName + packtype))
                {
                    var symFileNamewithoutExtension = Path.GetFileNameWithoutExtension(symFileName);

                    foreach (string newGen in newgenerationFileList)
                    {
                        rdbFileFoundLookup[symfile] = false;
                        if (newGen.ToLower().Contains(".pdf") && newGen.Trim().Contains(packtype.Trim()) && newGen.ToLower().Trim().Contains(clientName.ToLower().Trim()))
                        {
                            Console.WriteLine(newGen);
                            var newGenFileName = newGen;
                            var newGenFileNamewithoutExtension = Path.GetFileNameWithoutExtension(newGenFileName);
                            client.Add(clientName + packtype);
                            helper.writeToTextFile(resultPath + "\\processingFile.txt", clientName + packtype);
                            asposeHelper.convertPDFExcel(symFileName, Path.Combine(symphonyPath, symFileNamewithoutExtension) + ".xls");
                            asposeHelper.convertPDFExcel(newGenFileName, Path.Combine(newGenerationPath, newGenFileNamewithoutExtension) + ".xls");
                            var oldInvoiceReport = valuationHelper.getSymphonyInvoiceDetails(Path.Combine(symphonyPath, symFileNamewithoutExtension) + ".xls", clientName, packtype);
                            var newInvoiceReport = valuationHelper.getInvoiceDetails(Path.Combine(newGenerationPath, newGenFileNamewithoutExtension) + ".xls", clientName, packtype);
                            var invoiceResult = valuationHelper.compareInvoice(oldInvoiceReport, newInvoiceReport);
                            var oldPerformanceReport = valuationHelper.getSymphonyPortfolioPerformance(Path.Combine(symphonyPath, symFileNamewithoutExtension) + ".xls", clientName);
                            var newPerformanceReport = valuationHelper.getPortfolioPerformance(Path.Combine(newGenerationPath, newGenFileNamewithoutExtension) + ".xls", clientName, oldPerformanceReport);
                            var performanceResult = valuationHelper.comparePerformanceReport(oldPerformanceReport, newPerformanceReport);
                            var oldValuationReport = valuationHelper.getSymphonyQuarterlyValuationData(Path.Combine(symphonyPath, symFileNamewithoutExtension) + ".xls", clientName);
                            var newValuationReport = valuationHelper.getNewReportQuarterlyValuationData(Path.Combine(newGenerationPath, newGenFileNamewithoutExtension) + ".xls", clientName);
                            var oldCashReport = monthlyhelper.getOldReportCashStatements(Path.Combine(symphonyPath, symFileNamewithoutExtension) + ".xls", clientName);
                            var newCashReport = monthlyhelper.getNewReportCashStatements(Path.Combine(newGenerationPath, newGenFileNamewithoutExtension) + ".xls", clientName);
                            var oldAqusitionReport = valuationHelper.getSymphonyQuarterlytAquisitionData(Path.Combine(symphonyPath, symFileNamewithoutExtension) + ".xls", clientName);
                            var newAqusitionReport = valuationHelper.getNewReportQuarterlyAquisitionData(Path.Combine(newGenerationPath, newGenFileNamewithoutExtension) + ".xls", clientName);
                            var valuationResult = valuationHelper.compareQuarterlyValuationReports(oldValuationReport, newValuationReport);
                            var cashResult = monthlyhelper.compareCashStatements(oldCashReport, newCashReport);
                            var acquisitionResult = valuationHelper.compareQuarterlyAqusitionData(oldAqusitionReport, newAqusitionReport);
                            var reconcileMarketValues = valuationHelper.reconcileTotalValues();
                            
                            valuationHelper.createJsonQuarterlyReport(valuationResult, cashResult, acquisitionResult, performanceResult, invoiceResult, reconcileMarketValues, Path.Combine(resultPath, "jsonResults", symFileNamewithoutExtension) + ".json");
                            rdbFileFoundLookup[symfile] = true;
                            break;
                        }
                    }
                }
                

            }

            var missingFiles = rdbFileFoundLookup.Where(r => !r.Value).Select(r => r.Key).ToList();
            File.Delete(missingRdbReportsFile);
            File.AppendAllLines(missingRdbReportsFile, missingFiles);

            Directory.CreateDirectory(Path.Combine(resultPath, "csvResult"));
            var csv = new StringBuilder();
            var first = "FileName";
            var second = "Section";
            var third = "description";
            var fourth = "issuedescription";
            var newLine = string.Format("{0},{1},{2},{3},{4}", first, second, third, fourth, "deviation");
            csv.AppendLine(newLine);
            File.WriteAllText(resultPath + "\\csvResult\\quarterly.csv", csv.ToString());
            valuationHelper.quarterlyReportTocsv(resultPath + "\\jsonResults", resultPath + "\\csvResult\\quarterly.csv");

        }
    }
    public void createJsonQuarterlyReportInvoice(List<invoiceproperty> invoiceList, string outPutPath)
    {
        staticholdingdescription = null;
        if (invoiceList.Count > 0)
        {
            List<quarterlyStruct> quarterlyList = new List<quarterlyStruct>();
            foreach (invoiceproperty invoiceItem in invoiceList)
            {
                quarterlyStruct quarterly = new quarterlyStruct();
                quarterly.client = invoiceItem.clientName;
                quarterly.invoice = invoiceItem;
                quarterlyList.Add(quarterly);
            }
            if (quarterlyList.Count > 0)
            {
                string jsonresponse = JsonConvert.SerializeObject(quarterlyList);
                System.IO.File.WriteAllText(outPutPath, jsonresponse);
            }
        }
    }
    public void createJsonQuarterlyReportValuation(List<valuationproperty> valuationList, string outPutPath)
    {
        staticholdingdescription = null;
        if (valuationList.Count > 0)
        {
            List<quarterlyStruct> quarterlyList = new List<quarterlyStruct>();
            foreach (valuationproperty valuationItem in valuationList)
            {
                quarterlyStruct quarterly = new quarterlyStruct();
                quarterly.client = valuationItem.clientName;
                quarterly.valuationproperty = valuationItem;
                quarterlyList.Add(quarterly);
            }
            if (quarterlyList.Count > 0)
            {
                string jsonresponse = JsonConvert.SerializeObject(quarterlyList);
                System.IO.File.WriteAllText(outPutPath, jsonresponse);
            }
        }
    }
}