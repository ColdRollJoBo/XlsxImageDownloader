using Bytescout.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using static System.Console;



namespace XlsxImageDownloader
{
    class Program
    {
        static void Main(string[] args)
        {
            WriteLine("Program Running...");
            // The Spreadsheet Class is the top most level of the document.
            // The workbook is one step in and is mostly what is used to get actual data.
            Spreadsheet vendorDoc = new Spreadsheet();
            vendorDoc.LoadFromFile(@"C:\Users\jbojovic\Desktop\WebScraperTest\destaco1pageonly.xlsx");

            Workbook vendorProductCatalog = vendorDoc.Workbook;
            int numberOfSheetsInCatalog = vendorProductCatalog.Worksheets.Count;

            // Setting this up so users can pick what columns they would like to use for data and using numbers
            // instead of typing the full name.
            // Dictionary<int, string> userMenuOptions = new Dictionary<int, string>();

            foreach (string sheet in CollectionOfAllSheetNames(vendorProductCatalog, numberOfSheetsInCatalog))
            {

                CreateMainProductFolder(sheet);
                LoopDownRowsInCurrentSheetAndDownloadImagesToCorrectFolders(vendorProductCatalog, sheet);



                // short-circuting out of the foreach loop
                break;


            }

            WriteLine("Broke out of loop");






            //string topLevelImagesFolder = @"C:\Users\jbojovic\Desktop\WebScraperTest\Images";
            //SheetDataSections fileDataParts = new SheetDataSections();



            //    fileDataParts.ProductNames = new List<string>();
            //    fileDataParts.SubProductCategories = new List<string>();
            //    fileDataParts.Urls = new List<string>();


            //    WriteLine("Please provide the following information.\n Enter the row number that the headers begin at.");
            //    //The starting row value has a minus one to account for the zero base indexing that excel uses.
            //    fileDataParts.StartingRowForLooping = int.Parse(ReadLine()) - 1;
            //    WriteLine("Please enter the Column name where the product names are located.");
            //    fileDataParts.ProductNameColumn = ReadLine().ToLower().Replace(" ", "");
            //    WriteLine("Please enter the Column name where the sub-folder names are located.");
            //    fileDataParts.SubProductFolderColumn = ReadLine().ToLower().Replace(" ", "").Replace("-", "");
            //    WriteLine("Please enter the Column name where the URL's are located.");
            //    fileDataParts.UrlColumn = ReadLine().ToLower().Replace(" ", "");


            //    Spreadsheet document = new Spreadsheet();
            //    document.LoadFromFile(@"C:\Users\jbojovic\Desktop\WebScraperTest\destaco.xlsx");
            //    int numberOfSheets = document.Workbook.Worksheets.Count;




            //    foreach (string sheetName in CollectionOfAllSheetNames(document, numberOfSheets))
            //    {

            //        Worksheet currentSheet = document.Workbook.Worksheets.ByName(sheetName);

            //        // Getting the last row in the sheet to use as the stopping point in the for loop.
            //        int lastRowInSheet = currentSheet.Rows.LastFormatedRow;



            //        // These 3 methods add all the product names, sub folder names and urls to their corresponding lists
            //        LoopDownTheColumnAndAddDataToCorrespondingList(currentSheet, fileDataParts.ProductNameColumn,
            //            fileDataParts.ProductNameColumn, fileDataParts.ProductNames, fileDataParts.StartingRowForLooping, lastRowInSheet);
            //        WriteLine($"Product Names: {fileDataParts.ProductNames.Count}");


            //        LoopDownTheColumnAndAddDataToCorrespondingList(currentSheet, fileDataParts.SubProductFolderColumn,
            //           fileDataParts.SubProductFolderColumn, fileDataParts.SubProductCategories, fileDataParts.StartingRowForLooping, lastRowInSheet);
            //        WriteLine($"Product SubFolders: {fileDataParts.SubProductCategories.Count}");

            //        LoopDownTheColumnAndAddDataToCorrespondingList(currentSheet, fileDataParts.UrlColumn,
            //           fileDataParts.UrlColumn, fileDataParts.Urls, fileDataParts.StartingRowForLooping, lastRowInSheet);
            //        WriteLine($"Product Urls: {fileDataParts.Urls.Count}");


            //        if (fileDataParts.ProductNames.Count == fileDataParts.Urls.Count && fileDataParts.ProductNames.Count == fileDataParts.SubProductCategories.Count)
            //        {
            //            ImageDownloader(currentSheet, fileDataParts.SubProductCategories, fileDataParts.ProductNames, fileDataParts.Urls);
            //        }
            //        else
            //        {
            //            WriteLine($"Sheet: {sheetName} does not have equal parts in all 3 lists");
            //            Task t = ErrorLog($"\rSheet: {sheetName} does not have equal parts in all 3 lists\r");
            //        }


            //        // Clearing the lists to make sure all the rows line up with the correct information from each sheet. 
            //        fileDataParts.ProductNames.Clear();
            //        fileDataParts.SubProductCategories.Clear();
            //        fileDataParts.Urls.Clear();




            //    }








            //    ReadLine();





        }

        public static List<string> CollectionOfAllSheetNames(Workbook catalog, int sheets)
        {
            List<string> allSheetsWorkbook = new List<string>();
            // This is a Zero based index so the -1 is needed in the for loop.
            for (int i = 0; i <= sheets - 1; i++)
            {
                string sheet = catalog.Worksheets[i].Name;
                allSheetsWorkbook.Add(sheet);
                WriteLine(sheet);

            }
            return allSheetsWorkbook;
        }

        public static void CreateMainProductFolder(string folderName)
        {
            // Setting the folder that all the other folders will be created in when organizing the downloaded images.
            string topLevelImagesFolder = @"C:\Users\jbojovic\Desktop\WebScraperTest\Images";
            Directory.CreateDirectory(topLevelImagesFolder + "\\" + folderName);

        }

        public static void LoopDownRowsInCurrentSheetAndDownloadImagesToCorrectFolders(Workbook book, string sheet)
        {
            Worksheet currentSheet = book.Worksheets.ByName(sheet);
            List<string> rowInfo = new List<string>();

            //  'i' is starting at 1 because there is header row describing the data below it.
            for (int i = 1; i <= currentSheet.Rows.LastFormatedRow - 1; i++)
            {
                for (int j = 0; j <= currentSheet.Columns.LastFormatedColumn; j++)
                {

                    switch (j)
                    {
                        case 0:
                            // Product Name at index[0]
                            if (currentSheet.Cell(i, j).ValueAsString != "")
                            {
                                rowInfo.Add(currentSheet.Cell(i, j).ValueAsString);
                            }
                            break;

                        case 1:
                            // Product Sub-Category at index[1]
                            if (currentSheet.Cell(i, j).ValueAsString != "")
                            {
                                rowInfo.Add(currentSheet.Cell(i, j).ValueAsString);
                            }
                            break;
                        case 2:
                            // Product Download Url at index[2]
                            if (currentSheet.Cell(i, j).ValueAsString != "")
                            {
                                rowInfo.Add(currentSheet.Cell(i, j).ValueAsString);
                            }
                            break;
                    }
                }
                if (rowInfo.Count.Equals(3))
                {
                    WriteLine($"{rowInfo[0]} : {rowInfo[1]} : {rowInfo[2]} ");
                }
                rowInfo.Clear();
            }


        }

        public static void CreatingMenuForUserToPickColumnsToUseForData(Workbook book, string sheet, Dictionary<int, string> menuOptions)
        {
            Worksheet currentSheet = book.Worksheets.ByName(sheet);

            // Looping through the columns in the current sheet.
            for (int i = 0; i <= currentSheet.Columns.LastFormatedColumn; i++)
            {
                if (currentSheet.Cell(4, i).ValueAsString.ToLower().Contains("product") ||
                    currentSheet.Cell(4, i).ValueAsString.ToLower().Contains("url") ||
                    currentSheet.Cell(4, i).ValueAsString.ToLower().Contains("category"))
                {

                    menuOptions.Add(i, currentSheet.Cell(4, i).ValueAsString);
                }
            }

        }

        //public static void LoopDownTheColumnAndAddDataToCorrespondingList(Worksheet currentSheet, string HeaderTitle, string SubCategoryName, List<string> listData, int loopStartRow, int lastRowInSheet)
        //{
        //    // Looping through the columns in the current sheet.
        //    for (int i = 0; i <= currentSheet.Columns.LastFormatedColumn; i++)
        //    {
        //        // Getting the string value of the header column cells to then use in the if statements below. 
        //        // We only want to get the data from certian columns that were specified in the beginning. 
        //        HeaderTitle = currentSheet.Cell(loopStartRow, i).ValueAsString.ToLower().Replace(" ", "").Replace("-", "");

        //        // TODO: Account for headers not being typed exactly (ie: parent category vs parent-category)
        //        // Look into regexing this process
        //        if (HeaderTitle.Equals(SubCategoryName))
        //        {

        //            // Looping down the columns to add the data to their associated lists.
        //            // adding a +1 to pass the Header Row in the sheet.
        //            for (int j = loopStartRow + 1; j <= lastRowInSheet; j++)
        //            {
        //                if (currentSheet.Cell(j, i).ValueAsString != null && currentSheet.Cell(j, i).ValueAsString != "")
        //                {
        //                    listData.Add(currentSheet.Cell(j, i).ValueAsString);
        //                }

        //            }
        //        }
        //    }

        //}


        //private static void ImageDownloader(Worksheet sheet, List<string> subCategories, List<string> productNames, List<string> downloadUrls)
        //{


        //    // Setting the top level folder that all the sub-folders will be created in when organizing the downloaded images.
        //    string topLevelImagesFolder = @"C:\Users\jbojovic\Desktop\WebScraperTest\Images";



        //    try
        //    {
        //        // Limiting the amount of products to loop through. uncommment the for loop underneath for all products.
        //        //for (int i = 0; i <= 10; i++)
        //        for (int i = 0; i < productNames.Count; i++)
        //        {
        //            // Looping through the list of product categories to set the folder names for the nested loop. 
        //            string productTypeFolder = subCategories[i];

        //            // Limiting the amount of urls to loop through. uncommment the for loop underneath for all urls.
        //            //for (int j = i; j <= 10;)
        //            for (int j = i; j <= downloadUrls.Count;)
        //            {
        //                try
        //                {   // Checking to see if the subfolder already exists. If so then the product is added to the folder.
        //                    if (Directory.CreateDirectory(topLevelImagesFolder + "\\" + sheet.Name + "\\" + productTypeFolder).Exists)
        //                    {
        //                        string webAddress = downloadUrls[j];
        //                        WebClient Client = new WebClient();
        //                        Client.DownloadFile(webAddress, $@"{topLevelImagesFolder}\{sheet.Name}\{productTypeFolder}\{productNames[i]}.jpg");
        //                        //WriteLine($"Item: {prodNames[i]} | Link: {urls[j]}");

        //                    }
        //                    // This else statement will create the sub-folder if it does not exist then add the images to the folder with the product name. 
        //                    else
        //                    {
        //                        Directory.CreateDirectory(topLevelImagesFolder + "\\" + sheet.Name + "\\" + productTypeFolder);
        //                        string webAddress = downloadUrls[j];
        //                        WebClient Client = new WebClient();
        //                        Client.DownloadFile(webAddress, $@"{topLevelImagesFolder}\{sheet.Name}\{productTypeFolder}\{productNames[i]}.jpg");
        //                    }
        //                }
        //                catch (Exception ex)
        //                {
        //                    WriteLine($"An exception has been caught: {ex.Message}. It is on item {productNames[i]} and the link is {downloadUrls[j]}");
        //                    Task t = ErrorLog(DateTime.Now + "\r" + sheet.Name + "\r" + productTypeFolder + "\r" + productNames[i] + "\r" + downloadUrls[j] + "\r");

        //                    string webAddress = downloadUrls[j];


        //                }
        //                //Thread.Sleep(500);
        //                break;
        //            }
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        WriteLine(ex.Message);
        //        Task t = ErrorLog(DateTime.Now + " : " + ex.Message);
        //        return;
        //    }

        //    WriteLine("Files have been created");
        //}


        public static async Task ErrorLog(string errorLogData)
        {
            using StreamWriter file = new(@"C:\Users\jbojovic\Desktop\FailedUrls.txt", append: true);
            await file.WriteLineAsync(errorLogData);
        }
    }
}
