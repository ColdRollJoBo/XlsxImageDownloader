using Bytescout.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using static System.Console;



namespace XlsxImageDownloader
{
    class Program
    {
        static void Main(string[] args)
        {
            string topLevelImagesFolder = @"C:\Users\jbojovic\Desktop\WebScraperTest\Images";
            SheetDataSections fileDataParts = new SheetDataSections();
            fileDataParts.ProductNames = new List<string>();
            fileDataParts.SubProductCategories = new List<string>();
            fileDataParts.Urls = new List<string>();


            WriteLine("Please provide the following information.\n Enter the row number that the headers begin at.");
            // The starting row value has a minus one to account for the zero base indexing that excel uses. 
            fileDataParts.StartingRowForLooping = int.Parse(ReadLine()) - 1;
            WriteLine("Please enter the Column name where the product names are located.");
            fileDataParts.ProductNameColumn = ReadLine().ToLower().Replace(" ", "");
            WriteLine("Please enter the Column name where the sub-folder names are located.");
            fileDataParts.SubProductFolderColumn = ReadLine().ToLower().Replace(" ", "").Replace("-", "");
            WriteLine("Please enter the Column name where the URL's are located.");
            fileDataParts.UrlColumn = ReadLine().ToLower().Replace(" ", "");


            Spreadsheet document = new Spreadsheet();
            document.LoadFromFile(@"C:\Users\jbojovic\Desktop\WebScraperTest\destaco.xlsx");
            int numberOfSheets = document.Workbook.Worksheets.Count;



            foreach (string sheetName in CollectionOfAllSheetNames(document, numberOfSheets))
            {

                Worksheet currentSheet = document.Workbook.Worksheets.ByName(sheetName);

                // Getting the last row in the sheet to use as the stopping point in the for loop.
                int lastRowInSheet = currentSheet.Rows.LastFormatedRow;

                // These 3 methods add all the product names, sub folder names and urls to their corresponding lists
                LoopDownTheColumnAndAddDataToCorrespondingList(currentSheet, fileDataParts.ProductNameColumn,
                    fileDataParts.ProductNameColumn, fileDataParts.ProductNames, fileDataParts.StartingRowForLooping, lastRowInSheet);
                WriteLine($"Product Names: {fileDataParts.ProductNames.Count}");


                LoopDownTheColumnAndAddDataToCorrespondingList(currentSheet, fileDataParts.SubProductFolderColumn,
                   fileDataParts.SubProductFolderColumn, fileDataParts.SubProductCategories, fileDataParts.StartingRowForLooping, lastRowInSheet);
                WriteLine($"Product SubFolders: {fileDataParts.SubProductCategories.Count}");

                LoopDownTheColumnAndAddDataToCorrespondingList(currentSheet, fileDataParts.UrlColumn,
                   fileDataParts.UrlColumn, fileDataParts.Urls, fileDataParts.StartingRowForLooping, lastRowInSheet);
                WriteLine($"Product Urls: {fileDataParts.Urls.Count}");

                ImageDownloader(currentSheet, fileDataParts.SubProductCategories, fileDataParts.ProductNames, fileDataParts.Urls);


                // Clearing the lists to make sure all the rows line up with the correct information from each sheet. 
                fileDataParts.ProductNames.Clear();
                fileDataParts.SubProductCategories.Clear();
                fileDataParts.Urls.Clear();

                


            }








            ReadLine();





        }

        public static List<string> CollectionOfAllSheetNames(Spreadsheet document, int sheets)
        {
            List<string> allSheetsInDocument = new List<string>();

            for (int i = 0; i <= sheets - 1; i++)
            {
                string sheet = document.Workbook.Worksheets[i].Name;
                allSheetsInDocument.Add(sheet);
                WriteLine(sheet);

            }
            return allSheetsInDocument;
        }

        public static void LoopDownTheColumnAndAddDataToCorrespondingList(Worksheet currentSheet, string HeaderTitle, string listObjectName, List<string> listData, int loopStartRow, int lastRowInSheet)
        {
            // Looping through the columns in the current sheet.
            for (int i = 0; i <= currentSheet.Columns.LastFormatedColumn; i++)
            {
                // Getting the string value of the header column cells to then use in the if statements below. 
                // We only want to get the data from certian columns that were specified in the beginning. 
                HeaderTitle = currentSheet.Cell(loopStartRow, i).ValueAsString.ToLower().Replace(" ", "").Replace("-", "");

                // TODO Account for headers not being typed exactly (ie: parent category vs parent-category)
                // Look into regexing this process
                if (HeaderTitle.Equals(listObjectName))
                {

                    // Looping down the columns to add the data to their associated lists.
                    // adding a +1 to pass the Header Row in the sheet.
                    for (int j = loopStartRow + 1; j <= lastRowInSheet; j++)
                    {
                        if (currentSheet.Cell(j, i).ValueAsString != null && currentSheet.Cell(j, i).ValueAsString != "")
                        {
                            listData.Add(currentSheet.Cell(j, i).ValueAsString);
                        }

                    }
                }
            }

        }


        private static void ImageDownloader(Worksheet allSheetsInWorkbook, List<string> subCategories, List<string> productNames, List<string> downloadUrls)
        {


            // Setting the top level folder that all the sub-folders will be created in when organizing the downloaded images.
            string topLevelImagesFolder = @"C:\Users\jbojovic\Desktop\WebScraperTest\Images";


            try
            {
                for (int i = 0; i <= 10; i++)
                //for (int i = 0; i < prodNames.Count; i++)
                {
                    // Looping through the list of product categories to set the folder names for the nested loop. 
                    string productTypeFolder = subCategories[i];

                    for (int j = i; j <= 10;)
                    //for (int j = i; j <= urls.Count;)
                    {
                        try
                        {
                            if (Directory.CreateDirectory(topLevelImagesFolder + "\\" + productTypeFolder).Exists)
                            {
                                string webAddress = downloadUrls[j];
                                WebClient Client = new WebClient();
                                Client.DownloadFile(webAddress, $@"{topLevelImagesFolder}\{productTypeFolder}\{productNames[i]}.jpg");
                                //WriteLine($"Item: {prodNames[i]} | Link: {urls[j]}");

                            }
                            // This else statement will create the sub-folder if it does not exist then add the images to the folder with the product name. 
                            else
                            {
                                Directory.CreateDirectory(topLevelImagesFolder + "\\" + productTypeFolder);
                                string webAddress = downloadUrls[j];
                                WebClient Client = new WebClient();
                                Client.DownloadFile(webAddress, $@"{topLevelImagesFolder}\{productTypeFolder}\{productNames[i]}.jpg");
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteLine($"An exception has been caught: {ex.Message}. It is on item {productNames[i]} and the link is {downloadUrls[j]}");
                            Task t = ErrorLog(DateTime.Now + "\r" + productNames[i] + "\r" + downloadUrls[j] + "\r");
                        }
                        break;
                    }
                }

            }
            catch (Exception ex)
            {
                WriteLine(ex.Message);
                Task t = ErrorLog(DateTime.Now + " : " + ex.Message);
                return;
            }

            WriteLine("Files have been created");
        }


        public static async Task ErrorLog(string errorLogData)
        {
            using StreamWriter file = new(@"C:\Users\jbojovic\Desktop\FailedUrls.txt", append: true);
            await file.WriteLineAsync(errorLogData);
        }






    }
}
