using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Aspose.Cells;
using Aspose.Pdf;


namespace ValuationReport
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                if (args != null)
                {
                    const string pdfRecDir = @"\\ruffer.local\dfs\Shared\PDFRec\";

                    if (args[0].ToLower() == "distribute")
                    {
                        Console.WriteLine("Within if loop : distribute");
                        var folder = new DirectoryInfo(pdfRecDir);
                        var subDir = args[3].Equals("monthly", StringComparison.InvariantCultureIgnoreCase) ? "Monthly" : "Quarterly";
                        if (folder.Exists && folder.GetFileSystemInfos().Length != 0)
                        {
                            Console.WriteLine("Deleting files within " + subDir + "\n");
                            string[] dirToDel = Directory.GetDirectories(pdfRecDir + subDir);
                            foreach (var dir in dirToDel)
                            {
                                Directory.Delete(dir, true);
                            }
                            string[] fileList = Directory.GetFiles(pdfRecDir + subDir);
                            foreach (var file in fileList)
                            {
                                File.Delete(file);
                            }
                        }
                        Console.WriteLine("Distributing files for " + args[2] + "\n");
                        if (args.Length == 4)
                        {
                            DistributeAndCopyNewFiles(args[1], args[2], args[3]);
                        }
                        else if (args.Length == 5)
                        {
                            DistributeAndCopyNewFiles(args[1], args[2], args[3], args[4]);
                        }

                        return;
                    }
                    if (args[0].ToLower() == "archive-distribute")
                    {

                        if (args[3].ToLower() == "monthly")
                        {
                            Console.WriteLine("Moving the current folder - S drive PDFRec {0} to S drive PDFRec_Results_archive folder", args[3]);
                            ArchiveAndDistributeSymphonyFiles(args[1], args[2], args[3], args[4]);
                        }
                        else if (args[3].ToLower() == "quarterly")
                        {
                            Console.WriteLine("Moving the current folder - S drive PDFRec {0} to S drive PDFRec_Results_archive folder", args[3]);
                            ArchiveAndDistributeSymphonyFiles(args[1], args[2], args[3]);
                        }

                        return;
                    }
                    else if (args[0] == "consolidateMonthlyCSV")
                    {
                        Console.WriteLine("inside consolidate CSV");
                        createConsolidatedCSV(args[1]);
                        return;
                    }

                    else if (args.Count() == 3)
                    {
                        AsposHelper asposeHelper = new AsposHelper();
                        var parentFolderPath = args[0];
                        var newGenerationPath = args[1];
                        var reportType = args[2];
                        var symphonyFormatDir = parentFolderPath + "\\symphonyFormat";
                        if (!Directory.Exists(symphonyFormatDir))
                        {
                            Directory.CreateDirectory(symphonyFormatDir);
                        }

                        if (reportType.ToLower().Contains("monthly"))
                        {
                            MonthlyValuationHelper valuationHelper = new MonthlyValuationHelper();

                            valuationHelper.CompareMonthlyReport(symphonyFormatDir, newGenerationPath, parentFolderPath);
                        }
                        else if (reportType.ToLower().Contains("quarterly"))
                        {

                            quarterlyHelper quarterlyHelper = new quarterlyHelper();
                            quarterlyHelper.CompareQuarterlyReport(symphonyFormatDir, newGenerationPath, parentFolderPath);

                        }

                    }
                    else
                    {
                        throw new Exception("There is a mismatch in parameters passed please check!");
                    }
                }
            }
            catch (Exception e)
            {
               Console.WriteLine($"Exception caught: '{e}'");
               Console.WriteLine($"FATAL Error!!!");

               Environment.Exit(999);
            }
        }


        private static void ArchiveAndDistributeSymphonyFiles(string symphonyFolder, string newGenFolder, string valuationType, string searchPattern = null)
        {
            var folder = valuationType.Equals("monthly", StringComparison.InvariantCultureIgnoreCase) ? "Monthly" : "Quarterly";
            string distributedPath = @"\\ruffer.local\dfs\Shared\PDFRec\" + folder + @"\";


            //copy the existing folder to archived 

            if (!String.IsNullOrEmpty(valuationType) && !String.IsNullOrEmpty(symphonyFolder) && !String.IsNullOrEmpty(newGenFolder))
            {
                //moving the files from monthly/quaterly to PDF_archive folder
                string folderWithDate = DateTime.UtcNow.Date.ToString("d").Replace(@"/", "");
                string archiveDirectory = @"\\ruffer.local\dfs\Shared\PDFRec_Results_archive\" + folderWithDate + @"\" + valuationType;
                string[] distributedFolders = Directory.GetDirectories(distributedPath);

                if (Directory.Exists(archiveDirectory))
                {
                    string[] archivedFolders = Directory.GetDirectories(archiveDirectory);
                    foreach (var dir in archivedFolders)
                    {
                        Directory.Delete(dir, true);
                    }
                    string[] fileList = Directory.GetFiles(archiveDirectory);
                    foreach (var file in fileList)
                    {
                        File.Delete(file);
                    }
                    Directory.Delete(archiveDirectory);
                }
                if (distributedFolders.Length > 0)
                {

                    DirectoryCopy(distributedPath, archiveDirectory, true);
                    foreach (var dir in distributedFolders)
                    {
                        Directory.Delete(dir, true);
                    }
                    string[] fileList = Directory.GetFiles(distributedPath);
                    foreach (var file in fileList)
                    {
                        File.Delete(file);
                    }

                }


                DistributeAndCopyNewFiles(symphonyFolder, newGenFolder, valuationType, searchPattern);

            }
            else
            {
                throw new ArgumentException("Either of Valuation type or symphny folder or RDB folder are empty");
            }


        }


        private static void DistributeAndCopyNewFiles(string symphonyFolder, string newGenFolder, string valuationType, string searchPattern = null)
        {
            var folder = valuationType.Equals("monthly", StringComparison.InvariantCultureIgnoreCase) ? "Monthly" : "Quarterly";
            string distributedPath = @"\\ruffer.local\dfs\Shared\PDFRec\" + folder + @"\";
            var rdbFolder = distributedPath + "newGeneration";
            var symphFolder = @"\\ruffer.local\dfs\Shared\PDFRec\" + folder + @"\symphonyFormat";

            if (valuationType.Equals("monthly", StringComparison.InvariantCultureIgnoreCase))
            {
                string[] folders = { @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART1",
                                    @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART2",
                                    @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART3",
                                    @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART4",
                                    @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART5",
                                    @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART6",
                                    @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART7",
                                    @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART8",
                                    @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART9",
                                    @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART10",
                                    @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART11",
                                    @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART12",
                                    @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART13",
                                    @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART14",
                                    @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART15",
                                    @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART16",
                                    @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART17",
                                    @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART18",
                                    @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART19",
                                    @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\PART20"
                                };
                int counter = 0;


                //create a symphonyFormat folder which in each of the distributed Part and copy the file from the location(passed as parameter)
                foreach (var symphomyFile in Directory.EnumerateFiles(symphonyFolder))
                {
                    var folderIndex = counter % folders.Length;
                    var fileName = Path.GetFileName(symphomyFile);
                    System.IO.Directory.CreateDirectory(Path.Combine(folders[folderIndex], "symphonyFormat"));
                    var destFile = Path.Combine(folders[folderIndex], "symphonyFormat", fileName);
                    File.Copy(symphomyFile, destFile, true);
                    counter++;
                }
                int symphFileCount = counter;

                //Copy files from subfolders within MVAl BD 2 or MVAL BD 5 A /B / C /M to folder- newGeneration/{ respective subfolder same as RDS folder structure}

                DirectoryInfo dir = new DirectoryInfo(newGenFolder);
                DirectoryInfo[] dirs = dir.GetDirectories("*" + searchPattern + "*");
                foreach (DirectoryInfo mvalDir in dirs)
                {
                    string temppath = Path.Combine(newGenFolder, mvalDir.Name);
                    string rdbSubFolder = Path.Combine(rdbFolder, mvalDir.Name);
                    if (!Directory.Exists(rdbSubFolder))
                    {
                        Directory.CreateDirectory(Path.Combine(rdbFolder, mvalDir.Name));

                    }

                    if (Directory.GetDirectories(temppath).Count() > 0)
                    {
                        DirectoryInfo[] mvalDirs = mvalDir.GetDirectories();

                        foreach (DirectoryInfo subDir in mvalDirs)
                        {
                            string subDirPath = Path.Combine(temppath, subDir.Name);
                            DirectoryCopy(subDirPath, rdbSubFolder, true);
                        }
                    }
                    else
                    {
                        throw new DirectoryNotFoundException(
                    "Expected to have sub folders within the RDS folder: "
                    + temppath);
                    }

                }

            }

            else if (valuationType.Equals("quarterly", StringComparison.InvariantCultureIgnoreCase))
            {
                Console.WriteLine("Distributing files for quarterly reports");
                string[] folders =
                {                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART1",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART2",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART3",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART4",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART5",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART6",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART7",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART8",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART9",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART10",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART11",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART12",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART13",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART14",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART15",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART16",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART17",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART18",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART19",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART20",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART21",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART22",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART23",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART24",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART25",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART26",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART27",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART28",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART29",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART30",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART31",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART32",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART33",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART34",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART35",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART36",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART37",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART38",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART39",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART40",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART41",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART42",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART43",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART44",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART45",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART46",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART47",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART48",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART49",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART50",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART51",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART52",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART53",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART54",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART55",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART56",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART57",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART58",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART59",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART60",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART61",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART62",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART63",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART64",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART65",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART66",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART67",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART68",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART69",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART70",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART71",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART72",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART73",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART74",
                                @"\\ruffer.local\dfs\Shared\PDFRec\Quarterly\PART75"
            };
                int counter = 0;
                //copy symphony files to quarterly\PART*\symphonyFormat
                foreach (var symphomyFile in Directory.EnumerateFiles(symphonyFolder))
                {
                    var folderIndex = counter % folders.Length;
                    var fileName = Path.GetFileName(symphomyFile);
                    System.IO.Directory.CreateDirectory(Path.Combine(folders[folderIndex], "symphonyFormat"));
                    var destFile = Path.Combine(folders[folderIndex], "symphonyFormat", fileName);
                    File.Copy(symphomyFile, destFile, true);
                    counter++;
                }
                int symphFileCount = counter;
                //create a newGeneration folder which in each of the distributed Part and copy the file from the location(passed as parameter)
                foreach (var newGenFile in Directory.EnumerateFiles(newGenFolder))
                {
                    var folderIndex = counter % folders.Length;
                    var fileName = Path.GetFileName(newGenFile);
                    System.IO.Directory.CreateDirectory(rdbFolder);
                    //System.IO.Directory.CreateDirectory(Path.Combine(distributedPath, "newGeneration"));
                    var destFile = Path.Combine(distributedPath, "newGeneration", fileName);
                    File.Copy(newGenFile, destFile, true);
                    counter++;
                }
            }


        }

        private static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs, string searchPattern = null)
        {
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);
            DirectoryInfo[] dirs = string.IsNullOrEmpty(searchPattern) ? dir.GetDirectories() : dir.GetDirectories("*" + searchPattern + "*");


            // If the source directory does not exist, throw an exception.
            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            // If the destination directory does not exist, create it.
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }


            // Get the file contents of the directory to copy.
            FileInfo[] files = dir.GetFiles();

            foreach (FileInfo file in files)
            {
                // Create the path to the new copy of the file.
                string temppath = Path.Combine(destDirName, file.Name);

                // Copy the file.
                file.CopyTo(temppath, false);
            }

            // If copySubDirs is true, copy the subdirectories.
            if (copySubDirs)
            {

                foreach (DirectoryInfo subdir in dirs)
                {
                    // Create the subdirectory.
                    string temppath = Path.Combine(destDirName, subdir.Name);

                    // Copy the subdirectories.
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }

        }

        private static void createConsolidatedCSV(string type)
        {
            //@param type  can be monthly or quaterly
            List<resultsStructProperty> monthlyResults = new List<resultsStructProperty>();
            AsposHelper asposeHelper = new AsposHelper();
            //const string distributedPath = @"\\ruffer.local\dfs\Shared\PDFRec_Results_archive\27022019\monthly\";
            string driveLoc = @"\\ruffer.local\dfs\Shared\PDFRec\";
            string combinedResFolder = $"consolidated_results_{type}.csv";
            string distributedPath = driveLoc + type + @"\";
            //const string distributedPath = @"\\ruffer.local\dfs\Shared\PDFRec\Monthly\";
            string[] distributedFolders = Directory.GetDirectories(distributedPath, "PART*");
            var combinedResPath = Path.Combine(distributedPath, combinedResFolder);
            if (File.Exists(combinedResPath))
            {
                File.Delete(combinedResPath);
            }

            var csv = new StringBuilder();
            var newLine = string.Format("{0},{1},{2},{3},{4}", "FileName", "Section", "Description", "IssueDescription", "deviation");
            csv.AppendLine(newLine);
            File.WriteAllText(combinedResPath, csv.ToString());
            string[] csvFileList = null;
            if (!String.IsNullOrEmpty(distributedPath))
            {
                //break;

                foreach (var folder in distributedFolders)
                {

                    if ((Directory.GetDirectories(folder, "csv*")).Length == 0)
                    {

                        newLine = $"Unable to locate a {type} result file under folder: {folder}";
                        File.AppendAllText(combinedResPath, newLine + Environment.NewLine);
                    }
                    else
                    {
                        var csvResDir = Directory.GetDirectories(folder, "csv*")[0];
                        csvFileList = Directory.GetFiles(csvResDir.ToString());


                        foreach (var csvFile in csvFileList)
                        {
                            if (Path.GetExtension(csvFile.ToString()).Contains(".csv"))
                            {
                                int startRowPosition = 1;
                                int sheetIndex = asposeHelper.getWorksheetsCount(csvFile) - 1;
                                // get max rows from worksheet  
                                var tuplerowsColumn = asposeHelper.getRowsColumns(csvFile, sheetIndex);
                                var startColumnIndex = asposeHelper.getColumnIndexString(csvFile, sheetIndex, "FileName");
                                for (int cellIterator = startRowPosition; cellIterator < tuplerowsColumn.Item1; cellIterator++)
                                {
                                    var fileName = asposeHelper.getCellValue(csvFile, sheetIndex, cellIterator, startColumnIndex);
                                    var section = asposeHelper.getCellValue(csvFile, sheetIndex, cellIterator, startColumnIndex + 1);
                                    var description = asposeHelper.getCellValue(csvFile, sheetIndex, cellIterator, startColumnIndex + 2);
                                    var issueDescription = asposeHelper.getCellValue(csvFile, sheetIndex, cellIterator, startColumnIndex + 3);
                                    var deviation = asposeHelper.getCellValue(csvFile, sheetIndex, cellIterator, startColumnIndex + 4);
                                    var locallist = new resultsStructProperty();
                                    if (!string.IsNullOrEmpty(fileName))
                                    {
                                        locallist.fileName = fileName;
                                        locallist.section = section;
                                        locallist.description = description;
                                        locallist.issueDescription = issueDescription;
                                        locallist.deviation = deviation;
                                        //monthlyResults.Add(locallist);
                                        //monthlyResults.Insert(cellIterator - 1, locallist);
                                        newLine = string.Format("{0},{1},{2},{3},{4}", fileName, section, description, issueDescription, deviation);
                                        File.AppendAllText(combinedResPath, newLine + Environment.NewLine);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }


    }
}
