using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

class MonthlyValuationHelper
{
    AsposHelper asposeHelper = new AsposHelper();
    static string staticholdingdescription;
    static Boolean isValSummaryFound;
    static string portfolioCurrency;
    static string planName;
    static string statementPeriod;
    public static string[] labelHeaderArray = { "Cash statement", "Date", "Description", "Payment", "Reciepts", "Balance", "Cash statements" };
    public static string[] bidClientCodeList = { "ACT002", "ANZ001", "ARC005", "BAR086", "BOO004", "BRI013", "COM004", "DOC002",
            "DOC002C", "DUN045", "DUN045C", "ECF001", "EDM002", "INT002", "INT002C", "INT010", "LEE005", "NEI003", "NOR035C", "NUF001", "ROL001", "RUF001", "WES013" };
    static Boolean isBidClient;
    static Boolean isException;
    static string fileName;
    public static bool ignoreReporting;

    public List<valuationproperty> getOldReportValuationData(string workBookPath, string ClientCode)
    {
        staticholdingdescription = null;
        // column index B-1 , C-2, D-3, E-4,F-5,G-6,H-7,I-8
        List<valuationproperty> oldreport = new List<valuationproperty>();
        int workSheetsCount = asposeHelper.getWorksheetsCount(workBookPath);
        try
        {
            for (int workSheetIterator = 2; workSheetIterator < workSheetsCount; workSheetIterator++)
            {
                // check for valuation record in worksheet
                var doesValuationExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Valuation as at");
                if (doesValuationExists)
                {

                    // get max rows from worksheet
                    var tuplerowsColumn = asposeHelper.getRowsColumns(workBookPath, workSheetIterator);
                    //  get starting row position
                    var startRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Holding Description");
                    var holdingDescriptionColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Holding Description");
                    var costColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Cost");
                    var costRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Cost");
                    var portfolioCodeColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, ClientCode);
                    var portfolioRow = asposeHelper.getRowFromString(workBookPath, workSheetIterator, ClientCode);
                    var portfolioCode = asposeHelper.getCellValue(workBookPath, workSheetIterator, portfolioRow, portfolioCodeColumnIndex);
                    var portfolioCoulmnExists = false;
                    var portfolioCoulmn = asposeHelper.getCellValue(workBookPath, workSheetIterator, costRowPosistion, costColumnIndex - 2);
                    if (!string.IsNullOrEmpty(portfolioCoulmn))
                    {
                        if (portfolioCoulmn.Equals("Portfolio", StringComparison.InvariantCultureIgnoreCase))
                        {
                            portfolioCoulmnExists = true;
                        }
                    }
                    var getRow = asposeHelper.getCellValue(workBookPath, workSheetIterator, costRowPosistion, costColumnIndex - 1);
                    var columnExistBetween = false;
                    if (!string.IsNullOrEmpty(getRow))
                    {
                        if (getRow.Equals("Holding", StringComparison.InvariantCultureIgnoreCase))
                        {
                            if (portfolioCoulmnExists)
                            {
                                if (holdingDescriptionColumnIndex + 2 != costColumnIndex - 1)
                                {
                                    columnExistBetween = true;
                                }

                            }
                            else
                            {
                                if (holdingDescriptionColumnIndex + 1 != costColumnIndex - 1)
                                {
                                    columnExistBetween = true;
                                }
                            }
                        }
                    }
                    else
                    {
                        var lastCheck = asposeHelper.getCellValue(workBookPath, workSheetIterator, costRowPosistion + 3, costColumnIndex - 1);
                        decimal holdingCheckvalue;
                        if (Decimal.TryParse(lastCheck, out holdingCheckvalue))
                        {
                            if (portfolioCoulmnExists)
                            {
                                if (costColumnIndex - 1 != 3)
                                {
                                    columnExistBetween = true;
                                }
                            }
                            else
                            {
                                if (costColumnIndex - 1 != 2)
                                {
                                    columnExistBetween = true;
                                }
                            }
                        }
                        else
                        {

                        }

                    }

                    for (int cellIterator = startRowPosistion + 2; cellIterator < tuplerowsColumn.Item1; cellIterator++)
                    {
                        if (portfolioCoulmnExists)
                        {
                            string holdingdecription;
                            if (columnExistBetween)
                            {
                                var holdingSubString = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 3);
                                if (!string.IsNullOrEmpty(holdingSubString))
                                {
                                    holdingdecription = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 2) + " " + holdingSubString;
                                }
                                else
                                {
                                    holdingdecription = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 2);
                                }

                            }
                            else
                            {
                                holdingdecription = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 2);
                            }
                            var isStringPresent = asposeHelper.searchStringWorksheet(@"testdata\assestdataset.xls", 0, holdingdecription);
                            if (!string.IsNullOrEmpty(holdingdecription) && !isStringPresent && holdingdecription.ToLower().Trim() != "total value of securities and cash" && !holdingdecription.ToLower().Trim().Contains("exchange rates used") && !holdingdecription.Equals("securities", StringComparison.InvariantCultureIgnoreCase))
                            {
                                string percentAgeofTotalValue;
                                if (columnExistBetween)
                                {
                                    percentAgeofTotalValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 8);
                                }
                                else
                                {
                                    percentAgeofTotalValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 7);
                                }
                                if (!string.IsNullOrEmpty(percentAgeofTotalValue))
                                {
                                    string bookCost;
                                    if (columnExistBetween)
                                    {
                                        bookCost = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 5);
                                    }
                                    else
                                    {
                                        bookCost = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 4);
                                    }
                                    // string  marketPrice;
                                    // if(columnExistBetween)
                                    // {
                                    //     //marketPrice = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 6);
                                    // }else{
                                    //    // marketPrice = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 5);
                                    // }
                                    string marketValue;
                                    if (columnExistBetween)
                                    {
                                        marketValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 7);
                                    }
                                    else
                                    {
                                        marketValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 6);
                                    }
                                    string holding;
                                    if (columnExistBetween)
                                    {
                                        holding = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 4);
                                    }
                                    else
                                    {
                                        holding = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 3);
                                    }
                                    if (holding.StartsWith("(") && holding.EndsWith(")"))
                                    {
                                        holding = holding.Replace("(", "-").Replace(")", "");
                                    }

                                    string grossYield;
                                    if (columnExistBetween)
                                    {
                                        grossYield = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 9);
                                    }
                                    else
                                    {
                                        grossYield = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 10);
                                    }
                                    string estimatedGrossIncome;
                                    if (columnExistBetween)
                                    {
                                        estimatedGrossIncome = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 10);
                                    }
                                    else
                                    {
                                        estimatedGrossIncome = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 9);
                                    }
                                    var locallist = new valuationproperty();
                                    locallist.holdingdescription = holdingdecription;
                                    if (string.IsNullOrEmpty(holding))
                                    {
                                        locallist.holding = null;
                                    }
                                    else
                                    {
                                        locallist.holding = (decimal.Parse(holding)).ToString();
                                    }

                                    locallist.clientName = ClientCode;
                                    if (!String.IsNullOrEmpty(bookCost))
                                    {
                                        locallist.bookcost = bookCost;
                                    }
                                    // if (!String.IsNullOrEmpty(marketPrice))
                                    // {
                                    //     var withOutCurrencyCode = Regex.Replace(marketPrice, "[^0-9.]", "");
                                    //     var formattedPrice = decimal.Parse(withOutCurrencyCode);
                                    //     locallist.marketprice = Math.Round(formattedPrice, 2).ToString();

                                    // }
                                    // else
                                    // {
                                    //     locallist.marketprice = null;
                                    // }
                                    if (!String.IsNullOrEmpty(marketValue))
                                    {
                                        locallist.marketvalue = marketValue;
                                    }
                                    if (!String.IsNullOrEmpty(percentAgeofTotalValue))
                                    {
                                        //  locallist.percentageoftotalvalue = percentAgeofTotalValue;
                                    }
                                    if (!String.IsNullOrEmpty(grossYield))
                                    {
                                        locallist.grossyield = grossYield;
                                    }
                                    if (!String.IsNullOrEmpty(estimatedGrossIncome))
                                    {
                                        // Give 2 digit tolerance 
                                        if (estimatedGrossIncome.Length > 3)
                                        {
                                            locallist.estimatedgrossincome = estimatedGrossIncome.Substring(0, estimatedGrossIncome.Length - 2);
                                        }
                                        else
                                        {
                                            locallist.estimatedgrossincome = estimatedGrossIncome;
                                        }


                                    }
                                    oldreport.Add(locallist);
                                }
                            }

                        }
                        else
                        {

                            string holdingdecription;
                            if (columnExistBetween)
                            {
                                var holdingSubString = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 2);
                                if (!string.IsNullOrEmpty(holdingSubString))
                                {
                                    holdingdecription = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 1) + " " + holdingSubString;
                                }
                                else
                                {
                                    holdingdecription = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 1);
                                }

                            }
                            else
                            {
                                holdingdecription = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 1);
                            }
                            var isStringPresent = asposeHelper.searchStringWorksheet(@"testdata\assestdataset.xls", 0, holdingdecription);
                            if (!string.IsNullOrEmpty(holdingdecription) && !isStringPresent && holdingdecription.ToLower().Trim() != "total value of securities and cash" && !holdingdecription.ToLower().Trim().Contains("exchange rates used") && !holdingdecription.Equals("securities", StringComparison.InvariantCultureIgnoreCase))
                            {
                                string percentAgeofTotalValue;
                                if (columnExistBetween)
                                {
                                    percentAgeofTotalValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 7);
                                }
                                else
                                {
                                    percentAgeofTotalValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 6);
                                }
                                if (!string.IsNullOrEmpty(percentAgeofTotalValue))
                                {
                                    string bookCost;
                                    if (columnExistBetween)
                                    {
                                        bookCost = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 4);
                                    }
                                    else
                                    {
                                        bookCost = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 3);
                                    }
                                    // string  marketPrice;
                                    // if(columnExistBetween)
                                    // {
                                    //     marketPrice = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 5);
                                    // }else{
                                    //     marketPrice = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 4);
                                    // }
                                    string marketValue;
                                    if (columnExistBetween)
                                    {
                                        marketValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 6);
                                    }
                                    else
                                    {
                                        marketValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 5);
                                    }
                                    string holding;
                                    if (columnExistBetween)
                                    {
                                        holding = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 3);
                                    }
                                    else
                                    {
                                        holding = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 2);
                                    }

                                    if (holding.StartsWith("(") && holding.EndsWith(")"))
                                    {
                                        holding = holding.Replace("(", "-").Replace(")", "");
                                    }
                                    string grossYield;
                                    if (columnExistBetween)
                                    {
                                        grossYield = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 8);
                                    }
                                    else
                                    {
                                        grossYield = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 7);
                                    }
                                    string estimatedGrossIncome;
                                    if (columnExistBetween)
                                    {
                                        estimatedGrossIncome = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 9);
                                    }
                                    else
                                    {
                                        estimatedGrossIncome = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 8);
                                    }

                                    var locallist = new valuationproperty();
                                    locallist.holdingdescription = holdingdecription;
                                    locallist.portfolioCode = portfolioCode;
                                    if (string.IsNullOrEmpty(holding))
                                    {
                                        holding = "NoValueSet";
                                    }
                                    else if (holding.Contains(".") || holding.Contains(","))
                                    {
                                        holding = (decimal.Parse(holding)).ToString();
                                    }
                                    locallist.holding = holding;
                                    locallist.clientName = ClientCode;
                                    if (!String.IsNullOrEmpty(bookCost))
                                    {
                                        locallist.bookcost = bookCost;
                                    }
                                    // if (!String.IsNullOrEmpty(marketPrice))
                                    // {
                                    //     var withOutCurrencyCode = Regex.Replace(marketPrice, "[^0-9.]", "");
                                    //     var formattedPrice = decimal.Parse(withOutCurrencyCode);
                                    //     locallist.marketprice = Math.Round(formattedPrice, 2).ToString();

                                    // }
                                    // else
                                    // {
                                    //     locallist.marketprice = null;
                                    // }
                                    if (!String.IsNullOrEmpty(marketValue))
                                    {
                                        locallist.marketvalue = marketValue;
                                    }
                                    if (!String.IsNullOrEmpty(percentAgeofTotalValue))
                                    {
                                        //  locallist.percentageoftotalvalue = percentAgeofTotalValue;
                                    }
                                    if (!String.IsNullOrEmpty(grossYield))
                                    {
                                        locallist.grossyield = grossYield;
                                    }
                                    if (!String.IsNullOrEmpty(estimatedGrossIncome))
                                    {
                                        // Give 2 digit tolerance 
                                        if (estimatedGrossIncome.Length > 3)
                                        {
                                            estimatedGrossIncome = estimatedGrossIncome.Substring(0, estimatedGrossIncome.Length - 2);
                                        }

                                        if (estimatedGrossIncome.StartsWith("(") && estimatedGrossIncome.EndsWith(")"))
                                        {
                                            estimatedGrossIncome = estimatedGrossIncome.Replace("(", "-").Replace(")", "");
                                        }
                                        locallist.estimatedgrossincome = estimatedGrossIncome;
                                    }
                                    oldreport.Add(locallist);
                                }
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
        }
        catch
        {
            isException = true;
            fileName = workBookPath;
        }
        return oldreport;
    }
    public List<valuationproperty> getNewReportValuationData(string workBookPath, string ClientCode)
    {
        staticholdingdescription = null;
        // Holding Description - 1 , PortfolioCode - 2, Holding - 3, Market Price -4, Market Value - 5, Book Cost - 6, Percentage of total value -7, Gross yield -8, Estimated Gross Income -9
        List<valuationproperty> newReport = new List<valuationproperty>();
        int workSheetsCount = asposeHelper.getWorksheetsCount(workBookPath);
        //int workSheetsCount = 2;
        try
        {
            for (int workSheetIterator = 1; workSheetIterator < workSheetsCount; workSheetIterator++)
            {
                // check for valuation record
                var valuationExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Valuation");
                var portfolioExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Portfolio");
                var doesValuationExists = ((valuationExists && portfolioExists) && portfolioExists);
                //var doesValuationExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Book");
                if (doesValuationExists)
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
                        if (!string.IsNullOrEmpty(holdingdecription) && holdingdecription.ToLower().Trim() != "total" && !isStringPresent && !holdingdecription.ToLower().Trim().Contains("days accrued interest of") && holdingdecription.ToLower().Trim() != "exchange rates used:")
                        {
                            var percentAgeofTotalValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 7);
                            if (!string.IsNullOrEmpty(percentAgeofTotalValue))
                            {
                                staticholdingdescription = holdingdecription;
                                var bookCost = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 6);
                                var marketPrice = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 4);
                                var marketValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 5);
                                var holding = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 3);
                                var portfolioCode = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 2);
                                if (holding.StartsWith("(") && holding.EndsWith(")"))
                                {
                                    holding = holding.Replace("(", "-").Replace(")", "");
                                }
                                var grossYield = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 8);
                                var estimatedGrossIncome = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 9);
                                var locallist = new valuationproperty();
                                if (holdingdecription.StartsWith("(") && holdingdecription.EndsWith(")"))
                                {
                                    holdingdecription = holdingdecription.Replace("(", "-").Replace(")", "");
                                }
                                locallist.holdingdescription = holdingdecription;
                                locallist.portfolioCode = portfolioCode;
                                // Round off to 4 digits to match with symphont report
                                if (string.IsNullOrEmpty(holding))
                                {
                                    locallist.holding = null;
                                }
                                else
                                {
                                    if ((holding.Length - holding.IndexOf(".") - 1) > 4)
                                    {
                                        locallist.holding = Math.Round(decimal.Parse(holding), 4).ToString().Replace("(", "").Replace(")", "");
                                    }
                                    else
                                    {
                                        var r = (decimal.Parse(holding)).ToString();
                                        locallist.holding = r;
                                    }
                                }
                                locallist.clientName = ClientCode;
                                if (!String.IsNullOrEmpty(bookCost))
                                {
                                    if (bookCost.StartsWith("(") && bookCost.EndsWith(")"))
                                    {
                                        bookCost = bookCost.Replace("(", "-").Replace(")", "");
                                    }

                                    locallist.bookcost = bookCost;

                                }

                                if (!String.IsNullOrEmpty(marketValue))
                                {
                                    if (marketValue.StartsWith("(") && marketValue.EndsWith(")"))
                                    {
                                        marketValue = marketValue.Replace("(", "-").Replace(")", "");
                                    }

                                    locallist.marketvalue = marketValue;

                                }
                                if (!String.IsNullOrEmpty(percentAgeofTotalValue))
                                {
                                    // locallist.percentageoftotalvalue = percentAgeofTotalValue;
                                }
                                if (!String.IsNullOrEmpty(grossYield))
                                {
                                    if (grossYield.StartsWith("(") && grossYield.EndsWith(")"))
                                    {
                                        grossYield = grossYield.Replace("(", "-").Replace(")", "");
                                    }
                                    locallist.grossyield = grossYield;

                                }
                                if (!String.IsNullOrEmpty(estimatedGrossIncome))
                                {
                                    // Give 2 digit tolerance 
                                    if (estimatedGrossIncome.Length > 3)
                                    {
                                        locallist.estimatedgrossincome = estimatedGrossIncome.Substring(0, estimatedGrossIncome.Length - 2).Replace("(", "").Replace(")", "");
                                    }
                                    else
                                    {
                                        if (estimatedGrossIncome.StartsWith("(") && estimatedGrossIncome.EndsWith(")"))
                                        {
                                            estimatedGrossIncome = estimatedGrossIncome.Replace("(", "-").Replace(")", "");
                                        }
                                        locallist.estimatedgrossincome = estimatedGrossIncome;
                                    }
                                }
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
        }
        catch
        {
            isException = true;
            fileName = workBookPath;
        }
        return newReport;
    }
    public List<valuationSummaryproperty> getOldReportValuationSummaryData(string workBookPath, string ClientCode)
    {
        staticholdingdescription = null;
        // column index B-1 , C-2, D-3, E-4,F-5,G-6,H-7,I-8
        List<valuationSummaryproperty> oldValSumReport = new List<valuationSummaryproperty>();
        int workSheetsCount = asposeHelper.getWorksheetsCount(workBookPath);
        List<valuationSummaryproperty> templSumReport = new List<valuationSummaryproperty>();
        List<valuationSummaryproperty> templSumReport1 = new List<valuationSummaryproperty>();
        isValSummaryFound = false;
        try
        {
            for (int workSheetIterator = 0; workSheetIterator < workSheetsCount; workSheetIterator++)
            {
                // check for valuation record in worksheet

                var doesValuationExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Valuation summary as");

                if (doesValuationExists)
                {
                    //isValSummaryFound = true;
                    int cellIterator = 0;
                    int tempListCount = 0;
                    // get max rows from worksheet
                    var tuplerowsColumn = asposeHelper.getRowsColumns(workBookPath, workSheetIterator);
                    //  get starting row position
                    var startRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Sector Analysis");
                    var sectorColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Sector Analysis");
                    var costColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Cost");
                    var costRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Cost");
                    for (cellIterator = startRowPosistion + 2; cellIterator <= tuplerowsColumn.Item1; cellIterator++)
                    {

                        var sector = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, sectorColumnIndex);
                        var bookCost = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, costColumnIndex);

                        if (!string.IsNullOrEmpty(sector))
                        {
                            var isStringPresent = asposeHelper.searchStringWorksheet(@"testdata\assestdataset.xls", 0, sector);
                            var gatherData = false;
                            string marketValue = null;
                            string percentageSector = null;
                            string estimatedGrossIncome = null;
                            if (!isStringPresent && !sector.ToLower().Contains("total") && !string.IsNullOrEmpty(sector))
                            {
                                if (!string.IsNullOrEmpty(bookCost))
                                {
                                    gatherData = true;
                                }

                            }
                            else
                            {
                                if (!sector.ToLower().Contains("total") && !string.IsNullOrEmpty(sector))
                                {
                                    if (!string.IsNullOrEmpty(bookCost))
                                    {
                                        gatherData = true;
                                    }
                                }


                            }
                            if (gatherData)
                            {

                                if (!string.IsNullOrEmpty(bookCost))
                                {
                                    bookCost = bookCost.Replace(",", "");
                                }

                                marketValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, costColumnIndex + 1);
                                if (!string.IsNullOrEmpty(marketValue))
                                {
                                    marketValue = marketValue.Replace(",", "");
                                }

                                percentageSector = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, costColumnIndex + 2);
                                if (!string.IsNullOrEmpty(percentageSector))
                                {
                                    percentageSector = percentageSector.Replace(",", "");
                                }

                                estimatedGrossIncome = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, costColumnIndex + 3);
                                if (!string.IsNullOrEmpty(estimatedGrossIncome))
                                {
                                    estimatedGrossIncome = estimatedGrossIncome.Replace(",", "");
                                }


                                var tempList = new List<string> { bookCost, marketValue, percentageSector, estimatedGrossIncome };
                                var locallist = new valuationSummaryproperty();
                                if (!string.IsNullOrEmpty(estimatedGrossIncome) && !isValSummaryFound)
                                {
                                    staticholdingdescription = sector;
                                    locallist.sectorAnalysis = sector;
                                    locallist.clientName = ClientCode;
                                    locallist.bookCost = bookCost;
                                    locallist.marketValue = marketValue;
                                    locallist.percentageInSector = percentageSector;
                                    locallist.estimatedGrossIncome = estimatedGrossIncome;
                                    templSumReport.Add(locallist);

                                }
                                else if (!string.IsNullOrEmpty(estimatedGrossIncome) && isValSummaryFound)
                                {
                                    var tmpGroup = templSumReport.Where(x => x.sectorAnalysis == sector).Select(f => f);
                                    if (tmpGroup.Count() == 1)
                                    {
                                        var tmpGroupList = tmpGroup.ToList();
                                        var bookCostTmp = (bookCost.Contains(".") ? Decimal.Parse(bookCost) : Convert.ToInt32(bookCost));
                                        var listBookCost = (tmpGroupList[0].bookCost.Contains(".") ? Decimal.Parse(tmpGroupList[0].bookCost) : Convert.ToInt32(tmpGroupList[0].bookCost));
                                        var marketValueTmp = (marketValue.Contains(".") ? Decimal.Parse(marketValue) : Convert.ToInt32(marketValue));
                                        var listMarketValue = (tmpGroupList[0].marketValue.Contains(".") ? Decimal.Parse(tmpGroupList[0].marketValue) : Convert.ToInt32(tmpGroupList[0].marketValue));
                                        var estimatedGrossIncomeTmp = (estimatedGrossIncome.Contains(".") ? Decimal.Parse(estimatedGrossIncome) : Convert.ToInt32(estimatedGrossIncome));
                                        var listEstimatedGrossIncome = (tmpGroupList[0].estimatedGrossIncome.Contains(".") ? Decimal.Parse(tmpGroupList[0].estimatedGrossIncome) : Convert.ToInt32(tmpGroupList[0].estimatedGrossIncome));

                                        staticholdingdescription = sector;

                                        locallist.clientName = ClientCode;
                                        locallist.sectorAnalysis = sector;
                                        locallist.bookCost = (bookCostTmp + listBookCost).ToString();
                                        locallist.marketValue = (marketValueTmp + listMarketValue).ToString();
                                        locallist.estimatedGrossIncome = (estimatedGrossIncomeTmp + listEstimatedGrossIncome).ToString();
                                        locallist.percentageInSector = percentageSector;
                                        locallist.clientName = ClientCode;
                                        templSumReport1.Add(locallist);
                                    }
                                    else
                                    {
                                        staticholdingdescription = sector;

                                        locallist.clientName = ClientCode;
                                        locallist.sectorAnalysis = sector;
                                        locallist.bookCost = bookCost;
                                        locallist.marketValue = marketValue;
                                        locallist.estimatedGrossIncome = estimatedGrossIncome;
                                        locallist.percentageInSector = percentageSector;
                                        locallist.clientName = ClientCode;
                                        templSumReport1.Add(locallist);
                                    }

                                }
                                //if (cellIterator == tuplerowsColumn.Item1 - 3 && !isValSummaryFound)
                                //{
                                //    isValSummaryFound = true;
                                //}
                                //    else if (cellIterator == tuplerowsColumn.Item1 - 3 && isValSummaryFound)
                                //    {

                                //        locallist = new valuationSummaryproperty();
                                //        List<valuationSummaryproperty> templSumReport2 = new List<valuationSummaryproperty>();
                                //        templSumReport2 = templSumReport1;


                                //        foreach (var tempSum in templSumReport)
                                //        {
                                //            var bFound = false;
                                //            foreach ( var tempSum1 in templSumReport2)
                                //            {                                        
                                //                if (tempSum.sectorAnalysis == tempSum1.sectorAnalysis)
                                //                {

                                //                    bFound = true;
                                //                    break;
                                //                }

                                //            }
                                //           if(!bFound)
                                //            {
                                //                templSumReport1.Add(tempSum);
                                //            }



                                //        }
                                //        templSumReport = templSumReport1;
                                //        templSumReport1 = new  List<valuationSummaryproperty>();
                                //    }
                            }

                        }


                    }
                    if (!isValSummaryFound)
                    {
                        isValSummaryFound = true;
                    }
                    else if (isValSummaryFound && templSumReport1.Count() >= 1)
                    {

                        var locallist = new valuationSummaryproperty();
                        List<valuationSummaryproperty> templSumReport2 = new List<valuationSummaryproperty>();
                        templSumReport2 = templSumReport1;


                        foreach (var tempSum in templSumReport)
                        {
                            var bFound = false;
                            foreach (var tempSum1 in templSumReport2)
                            {
                                if (tempSum.sectorAnalysis == tempSum1.sectorAnalysis)
                                {

                                    bFound = true;
                                    break;
                                }

                            }
                            if (!bFound)
                            {
                                templSumReport1.Add(tempSum);
                            }



                        }
                        templSumReport = templSumReport1;
                        templSumReport1 = new List<valuationSummaryproperty>();
                    }

                }

            }

            if (isValSummaryFound && oldValSumReport.Count == 0)
            {
                oldValSumReport = templSumReport;
            }

            if (oldValSumReport.Count == 0 && isValSummaryFound)
            {
                valuationSummaryproperty locallist = new valuationSummaryproperty();
                locallist.clientName = ClientCode;
                locallist.sectorAnalysis = "Empty Valuation Summary section";
                oldValSumReport.Add(locallist);
            }

            else if (oldValSumReport.Count == 0)
            {
                valuationSummaryproperty locallist = new valuationSummaryproperty();
                locallist.clientName = ClientCode;
                locallist.sectorAnalysis = "Symphony Format Issue";
                oldValSumReport.Add(locallist);
            }
        }
        catch
        {
            isException = true;
            fileName = workBookPath;
        }
        return oldValSumReport;
    }
    public List<valuationSummaryproperty> getNewReportValuationSummaryData(string workBookPath, string ClientCode)
    {
        staticholdingdescription = null;
        // column index B-1 , C-2, D-3, E-4,F-5,G-6,H-7,I-8
        List<valuationSummaryproperty> newValSumReport = new List<valuationSummaryproperty>();
        int workSheetsCount = asposeHelper.getWorksheetsCount(workBookPath);
        var stopCollecting = false;
        try
        {
            for (int workSheetIterator = 1; workSheetIterator < 2; workSheetIterator++)
            {
                // check for valuation record in worksheet
                var doesValuationExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Valuation Summary");
                if (doesValuationExists)
                {
                    // get max rows from worksheet
                    var tuplerowsColumn = asposeHelper.getRowsColumns(workBookPath, workSheetIterator);
                    //  get starting row position
                    var startRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Sector analysis");
                    var sectorColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Sector analysis");

                    var costColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "cost");
                    var costRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "cost");
                    for (int cellIterator = startRowPosistion + 1; cellIterator < tuplerowsColumn.Item1; cellIterator++)
                    {
                        var sector = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, sectorColumnIndex);
                        var bookCost = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, costColumnIndex);
                        var estimatedGrossIncome = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, costColumnIndex + 3);
                        if (!string.IsNullOrEmpty(sector))
                        {
                            var isStringPresent = asposeHelper.searchStringWorksheet(@"testdata\assestdataset.xls", 0, sector);
                            var gatherData = false;
                            string marketValue = null;
                            string percentageSector = null;



                            if (!isStringPresent && !sector.ToLower().Contains("total") && !string.IsNullOrEmpty(sector))
                            {
                                if (!string.IsNullOrEmpty(bookCost) && !bookCost.ToLower().Contains("valuation summary") && !string.IsNullOrEmpty(estimatedGrossIncome) && !estimatedGrossIncome.ToLower().Contains("valuation summary"))
                                {
                                    gatherData = true;
                                }

                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(bookCost) && !string.IsNullOrEmpty(sector) && !sector.ToLower().Contains("total"))
                                {
                                    gatherData = true;
                                }
                            }
                            if (gatherData)
                            {
                                marketValue = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, costColumnIndex + 1);
                                percentageSector = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, costColumnIndex + 2);

                                var locallist = new valuationSummaryproperty();
                                if (!string.IsNullOrEmpty(estimatedGrossIncome))
                                {
                                    if (marketValue.StartsWith("(") && marketValue.EndsWith(")"))
                                    {
                                        marketValue = "-" + marketValue.Replace("(", "").Replace(")", "").Trim();
                                    }
                                    if (marketValue.Contains("."))
                                    {
                                        marketValue = Decimal.Parse(marketValue).ToString();
                                    }
                                    if (bookCost.StartsWith("(") && bookCost.EndsWith(")"))
                                    {
                                        bookCost = "-" + bookCost.Replace("(", "").Replace(")", "").Trim();
                                    }
                                    if (bookCost.Contains("."))
                                    {
                                        bookCost = Decimal.Parse(bookCost).ToString();
                                    }
                                    if (estimatedGrossIncome.StartsWith("(") && estimatedGrossIncome.EndsWith(")"))
                                    {
                                        estimatedGrossIncome = "-" + estimatedGrossIncome.Replace("(", "").Replace(")", "").Trim();
                                    }
                                    if (estimatedGrossIncome.Contains("."))
                                    {
                                        estimatedGrossIncome = Decimal.Parse(estimatedGrossIncome).ToString();
                                    }
                                    if (percentageSector.StartsWith("(") && percentageSector.EndsWith(")"))
                                    {
                                        percentageSector = "-" + percentageSector.Replace("(", "").Replace(")", "").Trim();
                                    }
                                    staticholdingdescription = sector;
                                    locallist.sectorAnalysis = sector;
                                    locallist.clientName = ClientCode;
                                    locallist.bookCost = bookCost;
                                    locallist.marketValue = marketValue;
                                    locallist.percentageInSector = percentageSector;
                                    locallist.estimatedGrossIncome = estimatedGrossIncome;
                                    newValSumReport.Add(locallist);
                                }
                            }
                            if (sector.ToLower().Contains("total"))
                            {
                                stopCollecting = true;
                            }


                        }


                    }

                }

            }
        }
        catch
        {
            isException = true;
            fileName = workBookPath;
        }
        return newValSumReport;
    }
    public List<acquisitionDisposals> getOldReportAcquisitionData(string workBookPath, string ClientCode)
    {
        staticholdingdescription = null;
        List<acquisitionDisposals> oldReport = new List<acquisitionDisposals>();
        int workSheetsCount = asposeHelper.getWorksheetsCount(workBookPath);
        try
        {
            for (int workSheetIterator = 0; workSheetIterator < workSheetsCount; workSheetIterator++)
            {
                var doesAcquisitionExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Acquisitions");
                var doesStockNameExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Stock Name");
                var additionalDetails = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Additional trade and execution information");
                if (doesAcquisitionExists && doesStockNameExists && !additionalDetails)
                {
                    // get max rows from worksheet
                    var tuplerowsColumn = asposeHelper.getRowsColumns(workBookPath, workSheetIterator);
                    //  get starting row position
                    var startRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Stock Name");
                    var sectorDescriptionIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Stock Name");
                    var tradeDateIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Trade");
                    var quantityIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Quantity");
                    var priceIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Price");
                    var fxRateIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "FX");
                    var commissionIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Broker");
                    var totalIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Consideration");
                    for (int cellIterator = startRowPosistion + 2; cellIterator < tuplerowsColumn.Item1; cellIterator++)
                    {
                        var description = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, sectorDescriptionIndex);
                        if (!String.IsNullOrEmpty(description) && !description.Equals("PURCHASES", StringComparison.InvariantCultureIgnoreCase) && !description.Equals("SALES", StringComparison.InvariantCultureIgnoreCase) && !description.ToLower().Contains("all transactions are expressed in the account currency"))
                        {

                            var date = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, tradeDateIndex);
                            DateTime dDate;

                            var notDateformat = false;
                            try
                            {


                                if (DateTime.TryParse(date, out dDate))
                                {
                                    String.Format("{0:d/MM/yyyy}", dDate);
                                }
                                else
                                {
                                    notDateformat = true; // <-- Control flow goes here
                                }
                            }
                            catch
                            {

                            }
                            if (!String.IsNullOrEmpty(date) && !notDateformat)
                            {
                                var locallist = new acquisitionDisposals();
                                locallist.securityDescription = description;
                                locallist.clientName = ClientCode;
                                staticholdingdescription = description;
                                string formattedDate = null;
                                try
                                {
                                    formattedDate = DateTime.ParseExact(date, "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd MMM yy", CultureInfo.InvariantCulture);
                                }
                                catch (Exception e)
                                {

                                }
                                if (!string.IsNullOrEmpty(formattedDate))
                                {
                                    locallist.date = formattedDate.TrimStart('0');
                                    var quantity = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, quantityIndex);
                                    if (!String.IsNullOrEmpty(quantity))
                                    {
                                        locallist.quantity = ((long)decimal.Parse(quantity)).ToString();
                                    }
                                    else
                                    {
                                        var settleDate = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, quantityIndex - 1);
                                        if (settleDate.Split(" ").Count() > 0)
                                        {
                                            var settleDateList = settleDate.Split(" ");
                                            quantity = settleDate.Split(" ").ToList().Where(x => x != null && x.IndexOf(",") >= 0).Select(f => f).FirstOrDefault();
                                            //trim the decimal places
                                            //quantity = quantity.Substring(0,quantity.IndexOf("."));
                                            locallist.quantity = ((long)decimal.Parse(quantity)).ToString();
                                        }
                                    }

                                    var price = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, priceIndex);
                                    if (!String.IsNullOrEmpty(price))
                                    {
                                        if (price.Contains("$"))
                                        {
                                            price = price.Replace("$", "usd");
                                        }
                                        else if (price.Contains("p"))
                                        {
                                            price = "gbp " + price.Replace("p", "").Trim();
                                        }
                                        else if (price.Contains("£"))
                                        {
                                            price = price.Replace("£", "gbp");
                                        }
                                        locallist.price = price;
                                    }
                                    var fxrate = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, fxRateIndex);
                                    if (!String.IsNullOrEmpty(fxrate))
                                    {
                                        locallist.fxrate = fxrate;

                                    }
                                    var commission = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, commissionIndex);
                                    if (!String.IsNullOrEmpty(commission))
                                    {
                                        locallist.commission = commission;

                                    }
                                    var total = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, totalIndex);
                                    if (!String.IsNullOrEmpty(total))
                                    {
                                        locallist.totalGBP = total;

                                    }
                                    oldReport.Add(locallist);
                                }


                            }
                            else
                            {

                                int index = oldReport.FindIndex(r => r.securityDescription == staticholdingdescription);
                                if (index != -1)
                                {
                                    if (Regex.IsMatch(description, @"^\d"))
                                    {
                                        if (char.IsLetter(oldReport[index].securityDescription[oldReport[index].securityDescription.Length - 1]))
                                        {
                                            oldReport[index].securityDescription = staticholdingdescription + " " + description;
                                        }
                                        else
                                        {

                                            oldReport[index].securityDescription = staticholdingdescription + description;
                                        }
                                    }
                                    else
                                    {
                                        oldReport[index].securityDescription = staticholdingdescription + " " + description;
                                    }
                                }

                            }

                        }

                    }


                }

            }
        }
        catch
        {
            isException = true;
            fileName = workBookPath;
        }
        return oldReport;
    }
    public List<acquisitionDisposals> getNewReportAcquisitionData(string workBookPath, string ClientCode)
    {
        staticholdingdescription = null;
        List<acquisitionDisposals> newReport = new List<acquisitionDisposals>();
        int workSheetsCount = asposeHelper.getWorksheetsCount(workBookPath);
        try
        {
            for (int workSheetIterator = 1; workSheetIterator < workSheetsCount; workSheetIterator++)
            {
                var doesAcquisitionExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "FX rate");
                if (doesAcquisitionExists)
                {
                    // get max rows from worksheet
                    var tuplerowsColumn = asposeHelper.getRowsColumns(workBookPath, workSheetIterator);
                    //  get starting row position
                    var startRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Trade date");
                for (int cellIterator = startRowPosistion + 1; cellIterator < tuplerowsColumn.Item1 - 1; cellIterator++)
                {
                    var description = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 2);
                    var acquiOrDisposalText = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 1);

                    if (acquiOrDisposalText.ToLower().Contains("additional transaction details"))
                    {
                        cellIterator = tuplerowsColumn.Item1;
                    }
                    else
                    {
                        if (!String.IsNullOrEmpty(description) && !description.ToLower().Contains("total consideration"))

                        {


                            var date = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 1);
                            if (!String.IsNullOrEmpty(date))
                            {
                                staticholdingdescription = description;
                                var locallist = new acquisitionDisposals();
                                locallist.securityDescription = description;
                                locallist.clientName = ClientCode;
                                locallist.date = date;
                                if (!description.Equals("Acquisitions", StringComparison.InvariantCultureIgnoreCase) || !description.Equals("Disposals", StringComparison.InvariantCultureIgnoreCase) || description.ToLower().Contains("commission"))
                                {
                                    
                                    var quantity = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 3);
                                    if (!String.IsNullOrEmpty(quantity))
                                    {
                                        locallist.quantity = ((long)decimal.Parse(quantity)).ToString();
                                    }

                                    var price = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 4);

                                    var fxrate = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 5);

                                    if (!String.IsNullOrEmpty(price))
                                    {
                                        if (Regex.IsMatch(price, "[a-z]", RegexOptions.IgnoreCase))
                                        {
                                            locallist.price = price.ToLower();
                                        }
                                        else
                                        {
                                            fxrate = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 4);
                                            locallist.fxrate = fxrate;
                                        }
                                        if (!String.IsNullOrEmpty(fxrate) && String.IsNullOrEmpty(locallist.fxrate))
                                        {
                                            locallist.fxrate = fxrate;
                                        }
                                        var commission = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 6);
                                        if (!String.IsNullOrEmpty(commission))
                                        {
                                            locallist.commission = commission;
                                        }
                                        var total = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 7);
                                        if (!String.IsNullOrEmpty(total))
                                        {
                                            locallist.totalGBP = total;
                                        }

                                    }
                                    else
                                    {
                                         price = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 5);

                                         fxrate = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 6);

                                        if (Regex.IsMatch(price, "[a-z]", RegexOptions.IgnoreCase))
                                        {
                                            locallist.price = price;
                                        }
                                        
                                        if (!String.IsNullOrEmpty(fxrate))
                                        {
                                            locallist.fxrate = fxrate;
                                        }
                                        var commission = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 7);
                                        if (!String.IsNullOrEmpty(commission))
                                        {
                                            locallist.commission = commission;
                                        }
                                        var total = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, 8);
                                        if (!String.IsNullOrEmpty(total))
                                        {
                                            locallist.totalGBP = total;
                                        }
                                    }
                                    
                                    
                                    newReport.Add(locallist);
                                }
                            }
                            else
                            {
                                int index = newReport.FindIndex(r => r.securityDescription == staticholdingdescription);
                                if (index != -1)
                                {
                                    if (Regex.IsMatch(description, @"^\d"))
                                    {
                                        if (char.IsLetter(newReport[index].securityDescription[newReport[index].securityDescription.Length - 1]))
                                        {
                                            newReport[index].securityDescription = staticholdingdescription + " " + description;
                                        }
                                        else
                                        {
                                            newReport[index].securityDescription = staticholdingdescription + description;
                                        }
                                    }
                                    else
                                    {
                                        newReport[index].securityDescription = staticholdingdescription + " " + description;
                                    }
                                }
                            }
                        }

                    }
                

                }


                }

            }
        }
        catch
        {
            isException = true;
            fileName = workBookPath;
        }
        return newReport;
    }

    public List<valuationproperty> compareValuationReports(List<valuationproperty> oldReport, List<valuationproperty> newReport)
    {
        staticholdingdescription = null;

        List<valuationproperty> missingList = new List<valuationproperty>();

        var holdingIssue = false;

        foreach (valuationproperty old in oldReport)
        {
            valuationproperty locallist = new valuationproperty();
            var recordExists = newReport.Any(x => x.holdingdescription.Equals(old.holdingdescription, StringComparison.InvariantCultureIgnoreCase) && x.portfolioCode.Equals(old.portfolioCode, StringComparison.InvariantCultureIgnoreCase));
            if (recordExists)
            {
                //var newReportGroup = newReport.Where(x => x.holdingdescription != null && x.holdingdescription.Equals(old.holdingdescription, StringComparison.InvariantCultureIgnoreCase) && x.clientName.Equals(old.clientName, StringComparison.InvariantCulture)).Select(f => f);
                var newReportGroup = newReport.Where(x => x.holdingdescription != null && x.holdingdescription.Equals(old.holdingdescription, StringComparison.InvariantCultureIgnoreCase) && x.marketvalue.Equals(old.marketvalue, StringComparison.InvariantCulture) && x.portfolioCode.Equals(old.portfolioCode, StringComparison.InvariantCulture)).Select(f => f);
                if (newReportGroup.Count() == 1)
                {
                    foreach (valuationproperty newItem in newReportGroup)
                    {

                       if (old.holding != newItem.holding)
                        {
                            if (newItem.holding.Contains("."))
                            {
                                //exception where symphony reports rounds holding value to 3 decimal places while RDS shows upto 5 decimal places
                                if (newItem.holding.Substring(newItem.holding.IndexOf(".") + 1).Length > 3)
                                {
                                    locallist.holding = "holding value in RDS reports have more than 3 decimal places";
                                }
                                else
                                {
                                    holdingIssue = true;
                                }
                            }
                            else
                            {
                                holdingIssue = true;
                            }
                        }
                        if (holdingIssue)
                        {
                            locallist.holding = old.holding + "|" + newItem.holding;
                        }

                        if (old.holding != newItem.holding || old.bookcost != newItem.bookcost || old.estimatedgrossincome != newItem.estimatedgrossincome || old.grossyield != newItem.grossyield || old.portfolioCode != newItem.portfolioCode || old.marketvalue != newItem.marketvalue || old.marketprice != newItem.marketprice)
                        {
                            locallist.holdingdescription = old.holdingdescription;
                            locallist.bookcost = old.bookcost + "|" + newItem.bookcost;
                            locallist.estimatedgrossincome = old.estimatedgrossincome + "|" + newItem.estimatedgrossincome;
                            locallist.grossyield = old.grossyield + "|" + newItem.grossyield;

                            locallist.marketvalue = old.marketvalue + "|" + newItem.marketvalue;
                            locallist.percentageoftotalvalue = old.percentageoftotalvalue + "|" + newItem.percentageoftotalvalue;
                            locallist.portfolioCode = old.portfolioCode + "|" + newItem.portfolioCode;
                            locallist.clientName = old.clientName;
                            missingList.Add(locallist);
                        }
                        break;
                    }
                }
            }
            else
            {
                locallist.holdingdescription = "no matching holding description found";
                locallist.portfolioCode = old.portfolioCode;
                locallist.clientName = old.clientName;
            }
          
          
        }
        return missingList;
    }
       
    public List<cashproperty> getOldReportCashStatements(string workBookPath, string ClientCode)
    {
        staticholdingdescription = null;
        portfolioCurrency = null;
        int startSheet = 0;
        // description - 1 , date - 2 , payments - 3 , receipts - 4, balance - 5
        List<cashproperty> oldCashReport = new List<cashproperty>();
        int workSheetsCount = asposeHelper.getWorksheetsCount(workBookPath);
        int ballanceCarriedOccurance = 0;
        int ballanceFwdOccurance = 0;
        //var accountTypeShifted = false;
        try
        {
            for (int workSheetIterator = 0; workSheetIterator < workSheetsCount; workSheetIterator++)
            {

                //determine the portfolio based currency which can be validated on valuation summary
                var setValSummaryPage = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Valuation") && asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Holding Description");

                if (setValSummaryPage)
                {
                    var startRowPosistion = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Book");
                    var startColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Book");
                    var currency = asposeHelper.getCellValue(workBookPath, workSheetIterator, startRowPosistion + 2, startColumnIndex);
                    portfolioCurrency = currency.Substring(currency.IndexOf("(") + 1).Replace(")", "");
                    startSheet = workSheetIterator;
                    break;
                }
            }
            for (int workSheetIterator = startSheet; workSheetIterator < workSheetsCount; workSheetIterator++)
            {

                // check for cash statements record in workSheets
                var cashSheetExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Payments");
                bool portfolioCurrencyExists = false;
                bool otherPortfolioCurExists = false;
                string portfolioCode = null;
                string cashAccountName = null;
                
                if (cashSheetExists)
                {
                    var paymentsColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Payments");
                    var startRowPosition = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Description");
                    var cashCurrency = asposeHelper.getCellValue(workBookPath, workSheetIterator, startRowPosition + 1, paymentsColumnIndex);
                    portfolioCurrencyExists = cashCurrency.Contains(portfolioCurrency);
                    var portfolioColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, ClientCode);
                    var portfolioRowPosition = asposeHelper.getRowFromString(workBookPath, workSheetIterator, ClientCode);
                    portfolioCode = asposeHelper.getCellValue(workBookPath, workSheetIterator, portfolioRowPosition, portfolioColumnIndex);
                    otherPortfolioCurExists = !string.IsNullOrEmpty(cashCurrency) && !cashCurrency.Contains(portfolioCurrency);                    
                    string cashAccountNameText = asposeHelper.getCellValue(workBookPath, workSheetIterator, 0, 2);
                    if(string.IsNullOrEmpty(cashAccountNameText))
                    {
                        cashAccountNameText = asposeHelper.getCellValue(workBookPath, workSheetIterator, 0, 1);
                    }
                    cashAccountName = cashAccountNameText.Substring(0, cashAccountNameText.IndexOf("account")+7);
                }

                if (cashSheetExists && portfolioCurrencyExists || cashSheetExists && otherPortfolioCurExists)
                {
                    
                    // get max rows from worksheet
                    var tuplerowsColumn = asposeHelper.getRowsColumns(workBookPath, workSheetIterator);
                    //  get starting row position
                    var balanceBroughtforwardExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Balance brought forward");
                    var balanceCarriedForwardExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Balance carried forward");
                    var startRowPosition = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Description");
                    var descriptionColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Description");
                    var dateColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Date");
                    var paymentsColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Payments");
                    var receiptsColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Receipts");
                    var balanceColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Balance");

                    for (int cellIterator = startRowPosition + 2; cellIterator < tuplerowsColumn.Item1; cellIterator++)
                    {
                        var locallist = new cashproperty();
                        var description = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, descriptionColumnIndex);
                        var date = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, dateColumnIndex);
                        var balance = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, balanceColumnIndex);
                        var reciepts = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, receiptsColumnIndex);

                        if (String.IsNullOrEmpty(balance))
                        {
                            balance = reciepts;
                            reciepts = null;
                        }

                        //if (!String.IsNullOrEmpty(description) && !description.Equals("Balance brought forward", StringComparison.InvariantCultureIgnoreCase) && !description.Contains("Page") && !description.Equals("Balance carried forward", StringComparison.InvariantCultureIgnoreCase))
                        if (!String.IsNullOrEmpty(description) && !description.Contains("Page"))
                            {
                            if (otherPortfolioCurExists || portfolioCurrencyExists)
                            {
                                locallist.description = description;
                                locallist.clientName = ClientCode;
                                locallist.date = date;
                                //locallist.balance = balance;
                                locallist.portfolioCode = portfolioCode;
                                
                                if(description.Equals("Balance brought forward", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    //ballanceFwdOccurance++;
                                    locallist.ballanceOccurance = ballanceFwdOccurance;
                                    locallist.accountName = cashAccountName;
                                    locallist.balance = balance;
                                }
                                else if (description.Equals("Balance carried forward", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    //ballanceCarriedOccurance++;
                                    locallist.ballanceOccurance = ballanceCarriedOccurance;
                                    locallist.accountName = cashAccountName;
                                    locallist.balance = balance;
                                }
                                oldCashReport.Add(locallist);
                            }

                        }
                    }
                }
                
            }
            if (portfolioCurrency.Contains("£", StringComparison.InvariantCultureIgnoreCase))
            {
                portfolioCurrency = "GBP";
            }
            else if (portfolioCurrency.Contains("$", StringComparison.InvariantCultureIgnoreCase))
            {
                portfolioCurrency = "USD";
            }

            else
            {
                Console.WriteLine("portfolio currency is displayed as displayed currency name");
                //throw new ArgumentNullException("Unable to convert currency symbol to currency name");
            }
        }
        catch
        {
            isException = true;
            fileName = workBookPath;
        }
        return oldCashReport;
    }
    public List<cashproperty> getNewReportCashStatements(string workBookPath, string ClientCode)
    {
        
        // Date - 1 , Description -2 , PaymentsGBP -3,ReceiptsGBP - 4, BalanceGBP - 5
        List<cashproperty> newCashReport = new List<cashproperty>();
        int workSheetsCount = asposeHelper.getWorksheetsCount(workBookPath);
        planName = null;
        statementPeriod = null;
        isException = false;
        fileName = null;
        int ballanceFwdOccurance = 0;
        int ballanceCarriedOccurance = 0;
        string cashAccountName = null;

        try
        {
            for (int workSheetIterator = 1; workSheetIterator < workSheetsCount; workSheetIterator++)
            {
                var isAppendixPage = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Appendix");
            var isContentsPage = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Contents");
            string portfolioCode = null;

                if (!isAppendixPage && !isContentsPage)
                {
                    // check for cash statements record in workSheets
                    var cashSheetExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Cash statement") || asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Cash statements");


                    if (cashSheetExists)
                    {
                        var paymentPortfolioExists = asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Payment " + portfolioCurrency) || asposeHelper.searchStringWorksheet(workBookPath, workSheetIterator, "Payments " + portfolioCurrency);
                        bool otherPortfolioCurrencyExists = false;
                        var paymentColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Payment");
                        var balanceColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Balance");
                        //var balanceRowPosition = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Balance " + portfolioCurrency);
                        var balanceRowPosition = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Balance");
                        var dateRowPosition = balanceRowPosition;
                        var portfolioColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, ClientCode);
                        var portfolioRowPosition = asposeHelper.getRowFromString(workBookPath, workSheetIterator, ClientCode);
                        var portfolioCodeText = asposeHelper.getCellValue(workBookPath, workSheetIterator, portfolioRowPosition, portfolioColumnIndex);
                        

                        try
                        {

                            portfolioCode = portfolioCodeText.Split(" ").Where(x => x.Contains(ClientCode)).Select(f => f).FirstOrDefault();
                        }
                        catch { }
                        if(string.IsNullOrEmpty(portfolioCode))
                        {

                            portfolioCode = portfolioCodeText.Substring(portfolioCodeText.IndexOf(ClientCode)).Trim();
                        }                        

                        var paymentRow = asposeHelper.getCellValue(workBookPath, workSheetIterator, dateRowPosition, paymentColumnIndex);
                        otherPortfolioCurrencyExists = asposeHelper.getCellValue(workBookPath, workSheetIterator, dateRowPosition, paymentColumnIndex).Length > 0;

                        if (cashSheetExists && paymentPortfolioExists || cashSheetExists && otherPortfolioCurrencyExists)
                        {
                            var tuplerowsColumn = asposeHelper.getRowsColumns(workBookPath, workSheetIterator);

                            var dateColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Date");
                            var descriptionColumnIndex = asposeHelper.getColumnIndexString(workBookPath, workSheetIterator, "Description");
                            var currency = paymentRow.Split(" ").ToList()[1];

                            var cashStatementRow = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Cash statement");

                            for (int i = cashStatementRow; i <= dateRowPosition; i++)
                            {
                                if (asposeHelper.getCellValue(workBookPath, workSheetIterator, i, dateColumnIndex).Contains(currency) && asposeHelper.getCellValue(workBookPath, workSheetIterator, i, dateColumnIndex).Contains("account"))
                                {
                                    statementPeriod = asposeHelper.getCellValue(workBookPath, workSheetIterator, i, dateColumnIndex);
                                    cashAccountName = statementPeriod.Substring(0, statementPeriod.IndexOf("account") + 7);
                                    break;
                                }
                            }
                            if (string.IsNullOrEmpty(planName))
                            {
                                planName = asposeHelper.getCellValue(workBookPath, workSheetIterator, cashStatementRow + 1, dateColumnIndex);
                                var list = planName.Split(" ").ToList();
                                if(list.Count > 0)
                                {
                                    if(list.Count > 5) 
                                    {
                                        planName = String.Format("{0} {1} {2} {3} {4}",list[0], list[1], list[2], list[3], list[4]);
                                        
                                    }
                                    else if (list.Count == 3)
                                    {
                                        planName = String.Format("{0} {1} {2}", list[0], list[1], list[2]);
                                        
                                    }
                                    else if (list.Count == 2)
                                    {
                                        planName = String.Format("{0} {1}", list[0], list[1]);
                                        
                                    }
                                }

                            }


                            var balanceBroughtForwardRow = asposeHelper.getRowFromString(workBookPath, workSheetIterator, "Balance brought forward");
                            var balanceBroughtForward = asposeHelper.getCellValue(workBookPath, workSheetIterator, balanceBroughtForwardRow, descriptionColumnIndex);
                            var balanceBroughtForwardBalance = asposeHelper.getCellValue(workBookPath, workSheetIterator, balanceBroughtForwardRow, balanceColumnIndex);
                            for (int cellIterator = dateRowPosition + 1; cellIterator < tuplerowsColumn.Item1; cellIterator++)
                            {
                                var locallist = new cashproperty();
                                bool isHeader = false;

                                var description = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, descriptionColumnIndex);
                                var date = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, dateColumnIndex);
                                var balance = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, balanceColumnIndex);
                                isHeader = (Array.IndexOf(labelHeaderArray, date) >= 0 || Array.IndexOf(labelHeaderArray, description) >= 0 || Array.IndexOf(labelHeaderArray, balance) >= 0);
                                var isValidDescription = false;
                                bool isStatementPeriod = false;
                                bool dateNotAPlanName = false;
                                bool dateNotOtherPlanName = false;

                                if (!string.IsNullOrEmpty(description))
                                {
                                    isValidDescription = !description.Equals(ClientCode) && !planName.Contains(description);
                                }
                                else if (!string.IsNullOrEmpty(date))
                                {
                                    if (!string.IsNullOrEmpty(planName.Split("-").ToList()[0].Trim()))
                                    {
                                        dateNotAPlanName = date.Contains(planName.Split("-").ToList()[0].Trim());
                                    }
                                    try
                                    {
                                        var planNameList = planName.Replace("-", "").Split(" ").ToList();
                                        var otherPlanNameList = date.Replace("-", "").Split(" ").ToList();
                                        var result = planNameList.Intersect(otherPlanNameList);

                                        dateNotOtherPlanName = result.FirstOrDefault().Length > 0;
                                    }
                                    catch { }
                                    isValidDescription = !date.Contains(planName) && !date.Contains(ClientCode) && !dateNotAPlanName && !dateNotOtherPlanName;
                                }
                                else if (!string.IsNullOrEmpty(balance))
                                {
                                    isValidDescription = !balance.Contains(ClientCode) && !balance.Contains(planName);
                                }

                                //isStatementPeriod = ((string.IsNullOrEmpty(statementPeriod) && date.Contains(statementPeriod.Substring(statementPeriod.IndexOf("to") + 3))));
                                if (!string.IsNullOrEmpty(balance) && balance.StartsWith("(") && balance.EndsWith(")"))
                                {
                                    balance = balance.Replace("(", "-").Replace(")", "").Trim();
                                }
                                if (string.IsNullOrEmpty(balance))
                                {
                                    var reciepts = balance = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, balanceColumnIndex - 1);
                                    var payments = balance = asposeHelper.getCellValue(workBookPath, workSheetIterator, cellIterator, balanceColumnIndex - 2);
                                    balance = !string.IsNullOrEmpty(reciepts) ? reciepts : payments;
                                }

                                if (!string.IsNullOrEmpty(date))
                                {
                                    var dateList = statementPeriod.Split("to").ToList();
                                    var statementEndDate = dateList[dateList.Count() - 1].Trim();
                                    try
                                    {


                                        isStatementPeriod = (!string.IsNullOrEmpty(statementPeriod) && date.Contains(statementEndDate));

                                    }
                                    catch
                                    {
                                    }


                                    //if (!isHeader && !String.IsNullOrEmpty(date) && isValidDescription && !isStatementPeriod && description != "Balance brought forward" && description != "Balance carried forward")
                                    if (!isHeader && !String.IsNullOrEmpty(date) && isValidDescription && !isStatementPeriod)
                                    {
                                        try
                                        {
                                            if (!string.IsNullOrEmpty(date))
                                            {
                                                DateTime dDate;
                                                dDate = DateTime.Parse(date);
                                                date = dDate.ToShortDateString();

                                            }
                                        }
                                        catch
                                        {

                                        }
                                      

                                    }
                                }
                                else if (!isHeader && isValidDescription && description.Equals("Balance brought forward", StringComparison.InvariantCultureIgnoreCase))
                                {                                    
                                    //ballanceFwdOccurance++;                                                                            
                                        
                                    locallist.balance = balance;                                    
                                    //locallist.ballanceOccurance = ballanceFwdOccurance;
                                    locallist.accountName = cashAccountName;                                    
                                }
                                else if (!isHeader && isValidDescription &&  description.Equals("Balance carried forward", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    //ballanceCarriedOccurance++;

                                    locallist.balance = balance;
                                    //locallist.ballanceOccurance = ballanceCarriedOccurance;
                                    locallist.accountName = cashAccountName;
                                }
                                locallist.description = description;
                                locallist.clientName = ClientCode;
                                locallist.date = date;
                                locallist.portfolioCode = portfolioCode;
                                newCashReport.Add(locallist);
                            }
                        }
                    }

                }
        }
        }
        catch
        {
            isException = true;
            fileName = workBookPath;
        }

        return newCashReport;
    }

    public List<valuationSummaryproperty> compareValuationSummary(List<valuationSummaryproperty> oldReport, List<valuationSummaryproperty> newReport)
    {
        staticholdingdescription = null;
        List<valuationSummaryproperty> missingList = new List<valuationSummaryproperty>();
        bool recordExists = false;
        foreach (valuationSummaryproperty oldItem in oldReport)
        {
            valuationSummaryproperty locallist = new valuationSummaryproperty();
            recordExists = newReport.Any(x => x.sectorAnalysis.Equals(oldItem.sectorAnalysis, StringComparison.InvariantCultureIgnoreCase) && x.clientName.Equals(oldItem.clientName, StringComparison.InvariantCultureIgnoreCase));
            if(!recordExists)
            {                
                foreach(valuationSummaryproperty newItem in newReport)
                {
                    if(newItem.sectorAnalysis.Contains("-"))
                    {
                        newItem.sectorAnalysis = newItem.sectorAnalysis.Replace("-", " ");
                        recordExists = newItem.sectorAnalysis.Equals(oldItem.sectorAnalysis);
                        break;
                    }
                }
            }
            if (recordExists)
            {
                var newReportGroup = newReport.Where(x => x.sectorAnalysis != null && x.sectorAnalysis.Equals(oldItem.sectorAnalysis, StringComparison.InvariantCultureIgnoreCase) && x.clientName.Equals(oldItem.clientName, StringComparison.InvariantCulture)).Select(f => f);
                if(newReportGroup.Count() ==1)
                {
                    foreach(valuationSummaryproperty compareItem in newReportGroup)
                    {
                        if(!oldItem.bookCost.Equals(compareItem.bookCost, StringComparison.InvariantCultureIgnoreCase) || !oldItem.marketValue.Equals(compareItem.marketValue, StringComparison.InvariantCultureIgnoreCase) || !oldItem.percentageInSector.Equals(compareItem.percentageInSector, StringComparison.InvariantCultureIgnoreCase) || !oldItem.estimatedGrossIncome.Equals(compareItem.estimatedGrossIncome, StringComparison.InvariantCultureIgnoreCase))
                        {
                            locallist.clientName = oldItem.clientName;
                            locallist.sectorAnalysis = oldItem.sectorAnalysis;
                            locallist.bookCost = oldItem.bookCost + "|" + compareItem.bookCost;
                            locallist.marketValue = oldItem.marketValue + "|" + compareItem.marketValue;
                            locallist.percentageInSector = oldItem.percentageInSector + "|" + compareItem.percentageInSector;
                            locallist.estimatedGrossIncome = oldItem.estimatedGrossIncome + "|" + compareItem.estimatedGrossIncome;
                            missingList.Add(locallist);

                        }

                    }
                }
            }
            else
            {
                if(oldItem.sectorAnalysis == "Empty Valuation Summary section")
                {
                    locallist.marketValue = "Valuation Summary section contains grand total of value 0";
                }
                else
                {
                    locallist.marketValue = "Unable to find a record for sector Analysis in new report- " + oldItem.sectorAnalysis;
                }

                locallist.clientName = oldItem.clientName;
                locallist.sectorAnalysis = oldItem.sectorAnalysis;                
                missingList.Add(locallist);
             }
        }
         return missingList;
    }

    public List<cashproperty> compareCashStatements(List<cashproperty> oldCashReport, List<cashproperty> newCashReport)
    {
        staticholdingdescription = null;
        List<cashproperty> missingList = new List<cashproperty>();
        exceptionStruct exceptionStruct = new exceptionStruct();

        if (isException)
        {
            cashproperty locallist = new cashproperty();
            locallist.description = "Exception caught and file wasnt processed";
            locallist.date = fileName;
            missingList.Add(locallist);
        }
        else
        {

            foreach (cashproperty old in oldCashReport)
            {
                cashproperty locallist = new cashproperty();
                
                var recordExists = newCashReport.Any(x => !string.IsNullOrEmpty(x.description) && x.description.Equals(old.description, StringComparison.InvariantCultureIgnoreCase) && x.date.Equals(old.date, StringComparison.InvariantCultureIgnoreCase));
                
                if (recordExists)
                {
                    if (old.description == "Balance brought forward" || old.description == "Balance carried forward")
                    {
                        var newCashGroup = newCashReport.Where(x => x.description != null && x.description.Equals(old.description, StringComparison.InvariantCultureIgnoreCase) && x.accountName.Equals(old.accountName, StringComparison.InvariantCulture) && x.ballanceOccurance == old.ballanceOccurance).Select(f => f);
                        if (newCashGroup.Count() == 1)
                        {
                            foreach (cashproperty newItem in newCashGroup)
                            {
                                if (old.balance != newItem.balance)
                                {
                                    locallist.description = old.description + "|" + newItem.description;
                                    locallist.date = old.date + "|" + newItem.date;
                                    locallist.balance = old.balance + "|" + newItem.balance;
                                    locallist.clientName = old.clientName;
                                    missingList.Add(locallist);
                                }
                            }
                        }
                        else
                        {
                            locallist.description = "no matching holding description found";
                            locallist.date = old.date;
                            locallist.clientName = old.clientName;
                        }
                    }
                    else
                    {
                        var newReportGroup = newCashReport.Where(x => x.description != null && x.description.Equals(old.description, StringComparison.InvariantCultureIgnoreCase) && x.date.Equals(old.date, StringComparison.InvariantCulture) && x.ballanceOccurance == old.ballanceOccurance).Select(f => f);
                        if (newReportGroup.Count() == 1)
                        {
                            foreach (cashproperty newItem in newReportGroup)
                            {
                                if (newItem.description == "Balance brought forward" || newItem.description == "Balance carried forward")
                                {
                                    if (old.balance != newItem.balance)
                                    {
                                        locallist.description = old.description + "|" + newItem.description;
                                        locallist.date = old.date + "|" + newItem.date;
                                        locallist.balance = old.balance + "|" + newItem.balance;
                                        locallist.clientName = old.clientName;
                                        missingList.Add(locallist);
                                    }
                                }

                                if (old.description != newItem.description || old.date != newItem.date || old.clientName != newItem.clientName)
                                {
                                    locallist.description = old.description + "|" + newItem.description;
                                    locallist.date = old.date + "|" + newItem.date;
                                    //locallist.balance = old.balance + "|" + newItem.balance;
                                    locallist.clientName = old.clientName;
                                    missingList.Add(locallist);
                                }
                                break;
                            }
                        }
                        else
                        {
                            locallist.description = "no matching holding description found";
                            locallist.date = old.date;
                            locallist.clientName = old.clientName;
                        }
                    }
                   


                }
            }
        }
        return missingList;
    }
    
    public void createJsonMonthlyValuation(List<valuationproperty> valuationList, List<cashproperty> cashList, List<acquisitionDisposals>AcquisitionList ,List<valuationSummaryproperty> valuationSummaryList ,string outPutPath)
    {
        staticholdingdescription = null;
        // Filter if there is no match in Client Name
        List<monthlyStruct> monthStructList = new List<monthlyStruct>();
        foreach (valuationproperty valuation in valuationList)
        {
            monthlyStruct monthly = new monthlyStruct();
            monthly.client = valuation.clientName;
            monthly.valuationproperty = valuation;
            monthStructList.Add(monthly);
        }
        foreach(cashproperty cashItem in cashList)
        {
                monthlyStruct monthly = new monthlyStruct();
                monthly.client = cashItem.clientName;
                monthly.cashproperty = cashItem;
                monthStructList.Add(monthly);
        }
        foreach(acquisitionDisposals acquistionItem in AcquisitionList)
        {
             monthlyStruct monthly = new monthlyStruct();
                monthly.client = acquistionItem.clientName;
                monthly.acquisitionDisposals = acquistionItem;
                monthStructList.Add(monthly);
        }
        foreach (valuationSummaryproperty vSummaryItem in valuationSummaryList)
        {
            monthlyStruct monthly = new monthlyStruct();
            monthly.client = vSummaryItem.clientName;
            monthly.valuationSummary = vSummaryItem;
            monthStructList.Add(monthly);
        }
        if (monthStructList.Count > 0)
        {
            string jsonresponse = JsonConvert.SerializeObject(monthStructList);
            System.IO.File.WriteAllText(outPutPath, jsonresponse);
        }

    }
    public void CompareMonthlyReport(string symphonyPath, string newGenerationPath, string resultPath)
    {
        staticholdingdescription = null;                
        Console.WriteLine("SymphonyPath - "+ symphonyPath);
        
        Helper commonhelper = new Helper();
        AsposHelper asposeHelper = new AsposHelper();
        MonthlyValuationHelper valuationHelper = new MonthlyValuationHelper();
        var isFile = commonhelper.checkFileOrDirectory(symphonyPath);
        List<bidClientException> bidClientExceptionList = new List<bidClientException>();
        if (isFile.Equals("directory", StringComparison.InvariantCultureIgnoreCase))
        {
            var symphonyFileList = commonhelper.getFileNames(symphonyPath, "*.pdf");
            //var newgenerationFileList = commonhelper.getFileNames(newGenerationPath, "*.pdf");
            List<string> client = new List<string>();
            
            int h = 0;
            foreach (string symfile in symphonyFileList)
            {
                isBidClient = false;
                h = h+1;
                Console.WriteLine("Total Files To Compare - " + symphonyFileList.Count());
                Console.WriteLine("Processing File - "+ h);
                string symFileName = symfile;
                var symFileLength = symFileName.Length;
                var lastIndex = symFileName.LastIndexOf("_");
                var clientName = symFileName.Substring(lastIndex + 1, symFileLength - (lastIndex + 1)).Replace(".pdf", "");
                var symFileNamewithoutExtension = Path.GetFileNameWithoutExtension(symFileName);
                string fileType = null;
                var localList = new bidClientException();
                Directory.CreateDirectory(Path.Combine(resultPath, "jsonResults"));
                if ( !client.Contains(clientName) )
                {
                    if(symFileName.Contains("MVAL BD5 M"))
                    {
                        fileType = "MVAL BD5 M";
                        if (Array.IndexOf(bidClientCodeList, clientName) >= 0)
                        {
                            isBidClient = true;
                            localList.clientName = clientName;
                            localList.isBidClient = isBidClient;
                            bidClientExceptionList.Add(localList);
                        }

                    } 
                    if(symFileName.Contains("MVAL BD5 A"))
                    {
                        fileType = "MVAL BD5 A";
                        if (Array.IndexOf(bidClientCodeList, clientName) >= 0)
                        {
                            isBidClient = true;
                            localList.clientName = clientName;
                            localList.isBidClient = isBidClient;
                            bidClientExceptionList.Add(localList);
                        }
                    } 
                    if(symFileName.Contains("MVAL BD5 B"))
                    {
                        fileType = "MVAL BD5 B";
                    } 
                    if(symFileName.Contains("MVAL BD5 C"))
                    {
                        fileType = "MVAL BD5 C";
                        if (Array.IndexOf(bidClientCodeList, clientName) >= 0)
                        {
                            isBidClient = true;
                            localList.clientName = clientName;
                            localList.isBidClient = isBidClient;
                            bidClientExceptionList.Add(localList);
                        }
                    } 
                     if(symFileName.Contains("M VAL BD2"))
                    {
                        fileType = "M VAL BD2";
                    }

                    var newgenerationFileList = commonhelper.getFileNames(newGenerationPath+@"\"+fileType, "*.pdf");
                    foreach (string newGen in newgenerationFileList)
                    {
                        if (newGen.ToLower().Contains(".pdf")  && newGen.Trim().Contains(fileType) && newGen.ToLower().Trim().Contains(clientName.ToLower().Trim()))
                        {
                            var newGenFileName = newGen;
                            var newGenFileNamewithoutExtension = Path.GetFileNameWithoutExtension(newGenFileName);
                            asposeHelper.convertPDFExcel(symFileName, Path.Combine(symphonyPath, symFileNamewithoutExtension) + ".xls");
                            asposeHelper.convertPDFExcel(newGenFileName, Path.Combine(newGenerationPath, newGenFileNamewithoutExtension) + ".xls");
                            var oldValuationSummaryReport = valuationHelper.getOldReportValuationSummaryData(Path.Combine(symphonyPath, symFileNamewithoutExtension) + ".xls", clientName);
                            var newValuationSummaryReport = valuationHelper.getNewReportValuationSummaryData(Path.Combine(newGenerationPath, newGenFileNamewithoutExtension) + ".xls", clientName);
                            var valuationSummaryResult = valuationHelper.compareValuationSummary(oldValuationSummaryReport, newValuationSummaryReport);
                            var oldValuationReport = valuationHelper.getOldReportValuationData(Path.Combine(symphonyPath, symFileNamewithoutExtension) + ".xls", clientName);
                            var newValuationReport = valuationHelper.getNewReportValuationData(Path.Combine(newGenerationPath, newGenFileNamewithoutExtension) + ".xls", clientName);
                            var valuationResult = valuationHelper.compareValuationReports(oldValuationReport, newValuationReport);
                            var oldCashReport = valuationHelper.getOldReportCashStatements(Path.Combine(symphonyPath, symFileNamewithoutExtension) + ".xls", clientName);
                            var newCashReport = valuationHelper.getNewReportCashStatements(Path.Combine(newGenerationPath, newGenFileNamewithoutExtension) + ".xls", clientName);
                            var cashResult = valuationHelper.compareCashStatements(oldCashReport, newCashReport);
                            var oldAcquisitionReport = valuationHelper.getOldReportAcquisitionData(Path.Combine(symphonyPath, symFileNamewithoutExtension) + ".xls", clientName);
                            var newAcquisitionReport = valuationHelper.getNewReportAcquisitionData(Path.Combine(newGenerationPath, newGenFileNamewithoutExtension) + ".xls", clientName);
                            var acquisitionresult = valuationHelper.compareAcquisitionData(oldAcquisitionReport, newAcquisitionReport);
                            if(!Directory.Exists(Path.Combine(resultPath, "jsonResults")))
                            {
                                Directory.CreateDirectory(Path.Combine(resultPath, "jsonResults"));
                            }
                            
                            valuationHelper.createJsonMonthlyValuation(valuationResult, cashResult,acquisitionresult, valuationSummaryResult, Path.Combine(resultPath, "jsonResults",symFileNamewithoutExtension) + ".json");
                            break;
                        }
                    }
                    client.Add(clientName);
                }
            }
            Directory.CreateDirectory(Path.Combine(resultPath, "csvResult"));
            var csv = new StringBuilder();
            var first = "FileName";
            var second = "Section";
            var thrid = "description";
            var fourth = "issuedescription";
            var newLine = string.Format("{0},{1},{2},{3},{4}", first, second,thrid,fourth,"deviation");
            csv.AppendLine(newLine);  
            File.WriteAllText(resultPath+"\\csvResult\\monthly.csv", csv.ToString());
            valuationHelper.monthlyReportwriteTocsv(resultPath+"\\jsonResults",resultPath+"\\csvResult\\monthly.csv", bidClientExceptionList);
        }
    }
    public List<acquisitionDisposals> compareAcquisitionData(List<acquisitionDisposals> oldAquReport, List<acquisitionDisposals> newAquReport)
    {
        staticholdingdescription = null;
        List<acquisitionDisposals> missingList = new List<acquisitionDisposals>();
        foreach (acquisitionDisposals oldItem in oldAquReport)
        {
            acquisitionDisposals locallist = new acquisitionDisposals();
            var recordExists = newAquReport.Any(x => x.securityDescription.Equals(oldItem.securityDescription, StringComparison.InvariantCultureIgnoreCase) && x.clientName.Equals(oldItem.clientName, StringComparison.InvariantCultureIgnoreCase));
            if (recordExists)
            {
                var oldReportGroup = oldAquReport.Where(x => x.date != null && x.date.Equals(oldItem.date, StringComparison.InvariantCultureIgnoreCase)  && (x.securityDescription != null && x.securityDescription.Equals(oldItem.securityDescription, StringComparison.InvariantCulture)));
                var newReportGroup = newAquReport.Where(x => x.date != null && x.date.Equals(oldItem.date, StringComparison.InvariantCultureIgnoreCase)  && (x.securityDescription!=null && x.securityDescription.Equals(oldItem.securityDescription, StringComparison.InvariantCulture)));
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
          
                    if (!oldReportGroup.FirstOrDefault().date.IsEqualOrBothNullOrEmpty(newReportGroup.FirstOrDefault().date) || !oldReportGroup.FirstOrDefault().price.IsEqualOrBothNullOrEmpty(newReportGroup.FirstOrDefault().price) || !oldReportGroup.FirstOrDefault().quantity.IsEqualOrBothNullOrEmpty(newReportGroup.FirstOrDefault().quantity))
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
                var finalCheckDescription = oldItem.securityDescription.Substring(0, (int)(oldItem.securityDescription.Length / 2));
                var descriptionExists = newAquReport.Any(x => x.securityDescription.Contains(finalCheckDescription) && x.clientName.Equals(oldItem.clientName, StringComparison.InvariantCultureIgnoreCase));
                if(!descriptionExists)
                {
                    locallist.securityDescription = oldItem.securityDescription;
                    locallist.date = "record missing in new generation report";
                    locallist.clientName = oldItem.clientName;
                    missingList.Add(locallist);
                }
            }
        }
        return missingList;
    }
    public void monthlyReportwriteTocsv(string folderPath, string outPutPath, List<bidClientException> bidClientExceptionList)
    {

        staticholdingdescription = null;
        Helper commonhelper = new Helper();
        var symphonyFileList = commonhelper.getFileNames(folderPath, "*.json");
        var filePath = outPutPath;
        var csv = new StringBuilder();
        foreach (string symfile in symphonyFileList)
        {
            List<monthlyStruct> monthlyList = JsonConvert.DeserializeObject<List<monthlyStruct>>(File.ReadAllText(symfile));
            var fileName = Path.GetFileNameWithoutExtension(symfile); ;
            foreach (monthlyStruct item in monthlyList)
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
                string AcquisitionDescription = null;
                string AcquisitionDate = null;
                string quantity = null;
                string price = null;
                string AcquisitionClientName = null;
                string fxRate = null;
                string vSummbookCost = null;
                string vSummsectorAnalysis = null;
                string vSumMarketValue = null;
                string vSummpercentageInSector = null;
                string vSummestimatedGrossIncome = null;
                string vSummclientName = null;
                var valuationExist = false;
                var valuationSummaryExist = false;
                var cashStatement = false;
                var AcquisitionExist = false;
                var symphonyPDFbad = false;
                var newReportAlignment = false;
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
                if (item.valuationSummary != null)
                {
                    vSummbookCost = item.valuationSummary.bookCost;
                    vSummclientName = item.valuationSummary.clientName;
                    vSummsectorAnalysis = item.valuationSummary.sectorAnalysis;
                    vSumMarketValue = item.valuationSummary.marketValue;
                    vSummpercentageInSector = item.valuationSummary.percentageInSector;
                    vSummestimatedGrossIncome = item.valuationSummary.estimatedGrossIncome;
                    valuationSummaryExist = true;
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
                    AcquisitionDescription = item.acquisitionDisposals.securityDescription;
                    AcquisitionDate = item.acquisitionDisposals.date;
                    quantity = item.acquisitionDisposals.quantity;
                    price = item.acquisitionDisposals.price;
                    AcquisitionClientName = item.acquisitionDisposals.clientName;
                    fxRate = item.acquisitionDisposals.fxrate;
                    AcquisitionExist = true;
                }
                // if (valuationExist)
                // {
                //     Helper common = new Helper();
                //     if (!string.IsNullOrEmpty(holding))
                //     {
                //         if (string.IsNullOrEmpty(common.splitString(holding, "|")[0]))
                //         {
                //             symphonyPDFbad = true;
                //         }
                //     }
                //     if (!string.IsNullOrEmpty(bookCost))
                //     {
                //         if (bookCost.Contains("|"))
                //         {
                //             if (string.IsNullOrEmpty(common.splitString(bookCost, "|")[0]))
                //             {
                //                 symphonyPDFbad = true;
                //             }
                //         }
                //     }
                    
                //     if (!string.IsNullOrEmpty(marketValue))
                //     {
                //         if (string.IsNullOrEmpty(common.splitString(marketValue, "|")[0]))
                //         {
                //             symphonyPDFbad = true;
                //         }
                //     }
                //     if (!string.IsNullOrEmpty(grossyield))
                //     {
                //         if (string.IsNullOrEmpty(common.splitString(grossyield, "|")[0]))
                //         {
                //             symphonyPDFbad = true;
                //         }
                //     }
                //     if (!string.IsNullOrEmpty(estimatedgrossincome))
                //     {
                //         if (string.IsNullOrEmpty(common.splitString(estimatedgrossincome, "|")[0]))
                //         {
                //             symphonyPDFbad = true;
                //         }
                //     }
                // }
                if (valuationExist)
                {
                    Helper common = new Helper();
                    if (!string.IsNullOrEmpty(holding))
                    {
                        if (!holding.Equals("holding value in RDS reports have more than 3 decimal places"))
                        {
                            if (string.IsNullOrEmpty(common.splitString(holding, "|")[1]) && !string.IsNullOrEmpty(common.splitString(holding, "|")[0]))
                            {
                                newReportAlignment = true;
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(bookCost))
                    {
                        if (bookCost.Contains("|"))
                        {
                            if (string.IsNullOrEmpty(common.splitString(bookCost, "|")[1]) && !string.IsNullOrEmpty(common.splitString(bookCost, "|")[0]))
                            {
                                newReportAlignment = true;
                            }
                        }
                    }
                    
                    if (!string.IsNullOrEmpty(marketValue))
                    {
                        if (string.IsNullOrEmpty(common.splitString(marketValue, "|")[1]) && !string.IsNullOrEmpty(common.splitString(marketValue, "|")[0]))
                        {
                            newReportAlignment = true;
                        }
                    }
                    if (!string.IsNullOrEmpty(grossyield))
                    {
                        if (string.IsNullOrEmpty(common.splitString(grossyield, "|")[1]) && !string.IsNullOrEmpty(common.splitString(grossyield, "|")[0]))
                        {
                            newReportAlignment = true;
                        }
                    }
                    if (!string.IsNullOrEmpty(estimatedgrossincome))
                    {
                        if (string.IsNullOrEmpty(common.splitString(estimatedgrossincome, "|")[1])&& !string.IsNullOrEmpty(common.splitString(estimatedgrossincome, "|")[0]))
                        {
                            newReportAlignment = true;
                        }
                    }
                }
                if (symphonyPDFbad)
                {
                    var section = "Alignment";
                    var description = "Tables header & rows are not aligned";
                    var issuedescription = "Symphony Report - Column Alignment Issue";
                    var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                    File.AppendAllText(filePath, newLine + Environment.NewLine);
                }
                if (newReportAlignment)
                {
                    var section = "Alignment";
                    var description = "Tables header & rows are not aligned";
                    var issuedescription = "New Generation Report - Column Alignment Issue";
                    var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                    File.AppendAllText(filePath, newLine + Environment.NewLine);
                }
                if (valuationExist && !symphonyPDFbad && !newReportAlignment)
                {
                    Helper common = new Helper();
                    if (!string.IsNullOrEmpty(holding) && !holding.Contains("NoValueSet"))
                    {
                        if(holding.Contains("|"))
                        {
                            ignoreReporting = false;
                            double oldReport;
                            double newReport;
                            Double.TryParse(common.splitString(holding, "|")[0], out oldReport);
                            Double.TryParse(common.splitString(holding, "|")[1], out newReport);

                            var difference = Math.Round(oldReport - newReport, 2);

                            if (difference.ToString().Length >= 1 && difference.GetType() == typeof(double))
                            {

                                ignoreReporting = difference < 0.1 || difference < -0.1;
                            }

                            
                            if (oldReport!= newReport && !ignoreReporting )
                            {
                                var section = "valuation";
                                var description = "holding value " + "(security decription: " + holdingDescription + " )"; 
                                var issuedescription = holding.Replace(",", "");
                                var deviation = difference.ToString();                               
                                var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription,deviation);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                            }
                        }else{
                            var section = "valuation";
                            var description = item.valuationproperty.holdingdescription.Replace(",", "");
                            var issuedescription = holding.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(bookCost))
                    {
                        if(bookCost.Contains("|"))
                        {
                            ignoreReporting = false;
                            double oldReport;
                            double newReport;
                            Double.TryParse(common.splitString(bookCost, "|")[0], out oldReport);
                            Double.TryParse(common.splitString(bookCost, "|")[1], out newReport);

                            var difference = Math.Round(oldReport - newReport, 2);

                            if (difference.ToString().Length > 1 && difference.GetType() == typeof(double))
                            {
                                ignoreReporting = difference < 0.1 || difference < -0.1;
                            }
                            
                            if (oldReport!= newReport && !ignoreReporting)
                            {
                                var section = "valuation";
                                var description = "bookCost value"; 
                                var issuedescription = bookCost.Replace(",", "");
                                var deviation = difference.ToString();
                                var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription,deviation);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                            }
                        }else{
                            var section = "valuation";
                            var description = item.valuationproperty.bookcost.Replace(",", "");
                            var issuedescription = bookCost.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(marketValue))
                    {
                        if(marketValue.Contains("|"))
                        {
                            ignoreReporting = false;
                            double oldReport;
                            double newReport;
                            Double.TryParse(common.splitString(marketValue, "|")[0], out oldReport);
                            Double.TryParse(common.splitString(marketValue, "|")[1], out newReport);
                            var difference = Math.Round(oldReport - newReport, 2);

                            if (difference.ToString().Length > 1 && difference.GetType() == typeof(double))
                            {
                                ignoreReporting = difference < 0.1 || difference < -0.1;
                            }

                            
                            if (oldReport!= newReport && !ignoreReporting)
                            {
                                var section = "valuation";
                                var description = "Market value"; 
                                var issuedescription = marketValue.Replace(",", "");
                                var deviation = difference.ToString();
                                var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription,deviation);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                            }
                        }else{
                            var section = "valuation";
                            var description = item.valuationproperty.marketvalue.Replace(",", "");
                            var issuedescription = marketValue.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(grossyield))
                    {
                        if(grossyield.Contains("|"))
                        {
                            double oldReport;
                            double newReport;
                            Double.TryParse(common.splitString(grossyield, "|")[0], out oldReport);
                            Double.TryParse(common.splitString(grossyield, "|")[1], out newReport);
                            var difference = oldReport - newReport;
                            
                            if (oldReport!= newReport)
                            {
                                if(difference > 1)
                                {
                                    var section = "valuation";
                                    var description = "gross yield value " + "( security decription: " + holdingDescription + " )";
                                    var issuedescription = grossyield.Replace(",", "");
                                    var deviation = Math.Round(oldReport - newReport, 2).ToString();
                                    var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription, deviation);
                                    File.AppendAllText(filePath, newLine + Environment.NewLine);
                                }
                                
                            }
                        }else{
                            var section = "valuation";
                            var description = item.valuationproperty.grossyield.Replace(",", "");
                            var issuedescription = grossyield.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    //commenting the validations upon having dicussion with fred as its not required(01/05/19)
                    //if (!string.IsNullOrEmpty(estimatedgrossincome))
                    //{
                    //    if(estimatedgrossincome.Contains("|"))
                    //    {
                    //        decimal oldReport;
                    //        decimal newReport;
                    //        Decimal.TryParse(common.splitString(estimatedgrossincome, "|")[0], out oldReport);
                    //        Decimal.TryParse(common.splitString(estimatedgrossincome, "|")[1], out newReport);
                    //        if(oldReport!= newReport)
                    //        {
                    //            var section = "valuation";
                    //            var description = "estimatedgrossincome value " + "(security decription: " + holdingDescription +" )"; 
                    //            var issuedescription = estimatedgrossincome.Replace(",", "");
                    //            var deviation = "Difference   " + (oldReport - newReport).ToString();
                    //            var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription,deviation);
                    //            File.AppendAllText(filePath, newLine + Environment.NewLine);
                    //        }
                    //    }else{
                    //        var section = "valuation";
                    //        var description = item.valuationproperty.estimatedgrossincome.Replace(",", "");
                    //        var issuedescription = estimatedgrossincome.Replace(",", "");
                    //        var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                    //        File.AppendAllText(filePath, newLine + Environment.NewLine);
                    //    }
                    //}
                }

                if (valuationSummaryExist)
                {
                    Helper common = new Helper();
                    bool bidClientRptIgnore = false;
                    var bidClientList = bidClientExceptionList.Where(x => x.clientName == vSummclientName).Select(f => f);
                    if (bidClientList.Count() == 1)
                    {
                        bidClientRptIgnore = bidClientList.ToList()[0].isBidClient;
                    }
                    if (!string.IsNullOrEmpty(vSumMarketValue))
                    {
                        
                        if (vSumMarketValue.Contains("|"))
                        {
                          
                           
                            ignoreReporting = false;
                            double oldReport;
                            double newReport;
                            Double.TryParse(common.splitString(vSumMarketValue, "|")[0].Replace("(", "-").Replace(")", ""), out oldReport);
                            Double.TryParse(common.splitString(vSumMarketValue, "|")[1].Replace("(", "-").Replace(")", ""), out newReport);
                            var difference = Math.Round(oldReport - newReport, 2);

                            if (difference.ToString().Length > 1 && difference.GetType() == typeof(double))
                            {
                                ignoreReporting = difference < 0.1 || difference < -0.1;
                            }

                                
                            if (oldReport != newReport && !ignoreReporting)
                            {

                                if (bidClientRptIgnore)
                                {
                                    var section = "valuation summary";
                                    var description = "market value";
                                    var issuedescription = item.valuationSummary.sectorAnalysis.Replace(",", "") + " ( difference: " + difference.ToString() + " )";
                                    var deviation = "Difference in market value to be ignored as this an Accepted BID Client(DR-1512)";
                                    var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription, deviation);
                                    File.AppendAllText(filePath, newLine + Environment.NewLine);
                                }
                                else
                                {


                                    var section = "valuation summary";
                                    var description = "market value";
                                    var issuedescription = item.valuationSummary.sectorAnalysis.Replace(",", "");
                                    var deviation = "Difference   " + Math.Round(difference, 2).ToString();
                                    var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription, deviation);
                                    File.AppendAllText(filePath, newLine + Environment.NewLine);
                                }
                            }
                            
                        }
                        else
                        {
                            var section = "valuation summary";
                            var description = item.valuationSummary.sectorAnalysis.Replace(",", "");
                            var issuedescription = vSumMarketValue.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    //commenting the validations upon having dicussion with fred as its not required(01/05/19)
                    //if (!string.IsNullOrEmpty(vSummpercentageInSector))
                    //{
                    //    if (vSummpercentageInSector.Contains("|"))
                    //    {
                    //        double oldReport;
                    //        double newReport;
                    //        Double.TryParse(common.splitString(vSummpercentageInSector, "|")[0].Replace("(", "-").Replace(")", ""), out oldReport);
                    //        Double.TryParse(common.splitString(vSummpercentageInSector, "|")[1].Replace("(", "-").Replace(")", ""), out newReport);
                    //        ignoreReporting = (oldReport - newReport) < 0.1;
                    //        if (oldReport != newReport && !ignoreReporting)
                    //        {
                    //            var section = "valuation summary";
                    //            var description = "Percentage Sector";
                    //            var issuedescription = item.valuationSummary.sectorAnalysis.Replace(",", "");
                    //            var deviation = "Difference   " + (oldReport - newReport).ToString();
                    //            var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription, deviation);
                    //            File.AppendAllText(filePath, newLine + Environment.NewLine);
                    //        }
                    //    }
                    //    else
                    //    {
                    //        var section = "valuation summary";
                    //        var description = item.valuationSummary.sectorAnalysis.Replace(",", "");
                    //        var issuedescription = vSummpercentageInSector.Replace(",", "");
                    //        var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                    //        File.AppendAllText(filePath, newLine + Environment.NewLine);
                    //    }
                    //}

                    //if (!string.IsNullOrEmpty(vSummestimatedGrossIncome))
                    //{
                    //    if (vSummestimatedGrossIncome.Contains("|"))
                    //    {
                    //        decimal oldReport;
                    //        decimal newReport;
                    //        Decimal.TryParse(common.splitString(vSummestimatedGrossIncome, "|")[0].Replace("(", "-").Replace(")", ""), out oldReport);
                    //        Decimal.TryParse(common.splitString(vSummestimatedGrossIncome, "|")[1].Replace("(", "-").Replace(")", ""), out newReport);
                    //        if (oldReport != newReport)
                    //        {
                    //            var section = "valuation summary";
                    //            var description = "Estimated Gross Income";
                    //            var issuedescription = item.valuationSummary.sectorAnalysis.Replace(",", "");
                    //            var deviation = "Difference   " + (oldReport - newReport).ToString();
                    //            var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription, deviation);
                    //            File.AppendAllText(filePath, newLine + Environment.NewLine);
                    //        }
                    //    }
                    //    else
                    //    {
                    //        var section = "valuation summary";
                    //        var description = item.valuationSummary.sectorAnalysis.Replace(",", "");
                    //        var issuedescription = vSummestimatedGrossIncome.Replace(",", "");
                    //        var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                    //        File.AppendAllText(filePath, newLine + Environment.NewLine);
                    //    }
                    //}

                    if (!string.IsNullOrEmpty(vSummbookCost))
                    {
                        if (vSummbookCost.Contains("|"))
                        {
                           
                            ignoreReporting = false;
                            double oldReport;
                            double newReport;
                            Double.TryParse(common.splitString(vSummbookCost, "|")[0].Replace("(", "-").Replace(")", ""), out oldReport);
                            Double.TryParse(common.splitString(vSummbookCost, "|")[1].Replace("(", "-").Replace(")", ""), out newReport);
                            var difference = Math.Round(oldReport - newReport, 2);

                            if (difference.ToString().Length > 1 && difference.GetType() == typeof(double))
                            {
                                ignoreReporting = difference < 0.1 || difference < -0.1;
                            }


                            if (oldReport != newReport && !ignoreReporting)
                            {
                                if (bidClientRptIgnore)
                                {
                                    var section = "valuation summary";
                                    var description = "book cost";
                                    var issuedescription = item.valuationSummary.sectorAnalysis.Replace(",", "") + " ( difference: " + difference.ToString() + " )";
                                    var deviation = "Difference in book cost to be ignored as this an Accepted BID Client(DR-1512)";
                                    var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription, deviation);
                                    File.AppendAllText(filePath, newLine + Environment.NewLine);
                                }
                                else
                                {
                                    var section = "valuation summary";
                                    var description = "Book Cost";
                                    var issuedescription = item.valuationSummary.sectorAnalysis.Replace(",", "");
                                    var deviation = "Difference   " + Math.Round(difference, 2).ToString();
                                    var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription, deviation);
                                    File.AppendAllText(filePath, newLine + Environment.NewLine);
                                }
                            }
                            
                        }
                        else
                        {
                            var section = "valuation summary";
                            var description = item.valuationSummary.sectorAnalysis.Replace(",", "");
                            var issuedescription = vSummbookCost.Replace(",", "");
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
                                var issuedescription = "cashDate value : " + common.splitString(cashDate, "|")[0].Replace(",", "") + "|" + common.splitString(cashDate, "|")[1].Replace(",", "");
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
                                var issuedescription = "paymentsreceipts value : " + common.splitString(paymentsreceipts, "|")[0].Replace(",", "") + "|" + common.splitString(paymentsreceipts, "|")[1].Replace(",", "");
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
                        if(cashBalance.Contains("|"))
                        {
                            if (!(common.splitString(cashBalance, "|")[0]).Equals(common.splitString(cashBalance, "|")[1], StringComparison.InvariantCultureIgnoreCase))
                            {
                                var section = "CashStatements";
                                var description = item.cashproperty.description.Replace(",", "");
                                var issuedescription = "cashBalance value : " + common.splitString(cashBalance, "|")[0].Replace(",", "") + "|" + common.splitString(cashBalance, "|")[1].Replace(",", "");
                                var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                            }
                        }else{
                                var section = "CashStatements";
                                var description = item.cashproperty.description.Replace(",", "");
                                var issuedescription = cashBalance.Replace(",", "");
                                var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                }
                if (AcquisitionExist && !symphonyPDFbad)
                {
                    Helper common = new Helper();
                    if (!string.IsNullOrEmpty(AcquisitionDate))
                    {
                        if (AcquisitionDate.Contains("|"))
                        {
                            if (!(common.splitString(AcquisitionDate, "|")[0]).Equals(common.splitString(AcquisitionDate, "|")[1], StringComparison.InvariantCultureIgnoreCase))
                            {
                                var section = "Acquisition";
                                var description = item.acquisitionDisposals.securityDescription.Replace(",", "");
                                var issuedescription = "acquisitionDate value : " + common.splitString(AcquisitionDate, "|")[0].Replace(",", "") + "|" + common.splitString(AcquisitionDate, "|")[1].Replace(",", "");
                                var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                                File.AppendAllText(filePath, newLine + Environment.NewLine);
                            }
                        }
                        else
                        {
                            var section = "Acquisition";
                            var description = item.acquisitionDisposals.securityDescription.Replace(",", "");
                            var issuedescription = AcquisitionDate.Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(quantity))
                    {
                        decimal oldQty;
                        decimal newQty;

                        var qtyValue1 = quantity.Split("|")[0];
                        var qtyValue2 = quantity.Split("|")[1];

                        Decimal.TryParse(qtyValue1, out oldQty);
                        Decimal.TryParse(qtyValue2, out newQty);
                        if (!(qtyValue1).Equals(qtyValue2, StringComparison.InvariantCultureIgnoreCase))
                        {
                            var section = "Acquisition";
                            var description = item.acquisitionDisposals.securityDescription.Replace(",", "");
                            var issuedescription = "quantity value : " + common.splitString(quantity, "|")[0].Replace(",", "") + "|" + common.splitString(quantity, "|")[1].Replace(",", "");
                            var deviation = "Difference in quantity value: " + (oldQty - newQty).ToString();
                            var newLine = string.Format("{0},{1},{2},{3}, {4}", fileName, section, description, issuedescription, deviation);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(price))
                    {
                        ignoreReporting = false;
                        double oldReportPrice;
                        double newReportPrice;

                        var priceValue1 = Regex.IsMatch(price.Split("|")[0], "[a-z]", RegexOptions.IgnoreCase) ? price.Split("|")[0].Split(" ")[1] : price.Split("|")[0];
                        var priceValue2 = Regex.IsMatch(price.Split("|")[1], "[a-z]", RegexOptions.IgnoreCase) ? price.Split("|")[1].Split(" ")[1] : price.Split("|")[1];

                        Double.TryParse(priceValue1, out oldReportPrice);
                        Double.TryParse(priceValue2, out newReportPrice);

                        var difference = Math.Round(oldReportPrice - newReportPrice, 2);

                        if (difference.ToString().Length > 1 && difference.GetType() == typeof(double))
                        {
                            ignoreReporting = difference < 0.1 || difference < -0.1;
                        }


                        if (!(oldReportPrice.ToString()).Equals(newReportPrice.ToString(), StringComparison.InvariantCultureIgnoreCase) && !ignoreReporting)
                        {
                            var section = "Acquisition";
                            var description = item.acquisitionDisposals.securityDescription.Replace(",", "");
                            var issuedescription = "price value : " + common.splitString(price, "|")[0].Replace(",", "") + "|" + common.splitString(price, "|")[1].Replace(",", "");
                            var deviation = "Difference in price value: " + difference.ToString();
                            var newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issuedescription, deviation);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                    if (!string.IsNullOrEmpty(fxRate))
                    {
                        if (!(common.splitString(fxRate, "|")[0]).Equals(common.splitString(fxRate, "|")[1], StringComparison.InvariantCultureIgnoreCase))
                        {
                            var section = "Acquisition";
                            var description = item.acquisitionDisposals.securityDescription.Replace(",", "");
                            var issuedescription = "fxRate value : " + common.splitString(fxRate, "|")[0].Replace(",", "") + "|" + common.splitString(fxRate, "|")[1].Replace(",", "");
                            var newLine = string.Format("{0},{1},{2},{3}", fileName, section, description, issuedescription);
                            File.AppendAllText(filePath, newLine + Environment.NewLine);
                        }
                    }
                }
                if(symphonyPDFbad)
                {
                    break;
                }
            }
        }
    }
   
}