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
        static bool triedAlready = false;
        static void Main(string[] args)
        {
            WriteLine("Program Running...");
            // The Spreadsheet Class is the whole excel file at the top most level.
            // The workbook is one step in and is mostly what is used to get actual data.
            Spreadsheet vendorDoc = new Spreadsheet();
            vendorDoc.LoadFromFile(@"C:\Users\jbojovic\Desktop\WebScraperTest\destaco3Column.xlsx");

            Workbook vendorProductCatalog = vendorDoc.Workbook;
            int numberOfSheetsInCatalog = vendorProductCatalog.Worksheets.Count;

            foreach (string sheet in CollectionOfAllSheetNames(vendorProductCatalog, numberOfSheetsInCatalog))
            {

                CreateMainProductFolder(sheet);

                LoopDownRowsInCurrentSheetAndDownloadImagesToCorrectFolders(vendorProductCatalog, sheet);

            }
            // For single Sheet Downloads by sheet name
            //LoopDownRowsInCurrentSheetAndDownloadImagesToCorrectFolders(vendorProductCatalog, "Conveyors");
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
            // Creating the new Sub-Category images will be organized into.
            Directory.CreateDirectory(topLevelImagesFolder + "\\" + folderName);

        }

        public static void LoopDownRowsInCurrentSheetAndDownloadImagesToCorrectFolders(Workbook book, string sheet)
        {
            Worksheet currentSheet = book.Worksheets.ByName(sheet);
            List<string> rowInfo = new List<string>();
            string mainImageFolder = @"C:\Users\jbojovic\Desktop\WebScraperTest\Images";

            //  'i' is starting at 1 because there is header row describing the data below it.
            for (int i = 1; i <= currentSheet.Rows.LastFormatedRow; i++)
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
                    DownloadImages(mainImageFolder, sheet, rowInfo);
                }
                else if (rowInfo.Count != 0)
                {
                    Task t = ErrorLog($"Row : {i} does not have all 3 parts needed to successfully download an item.");
                }
                rowInfo.Clear();
            }



        }


        public static void DownloadImages(string topLevelImagesFolder, string sheet, List<string> itemParts)
        {
            // Item parts consist of the name of the product, the sub-catagory and the url.


            try
            {   // Checking to see if the subfolder already exists. If so then the product is added to the folder.
                if (Directory.CreateDirectory(topLevelImagesFolder + "\\" + sheet + "\\" + itemParts[1]).Exists)
                {
                    string webAddress = itemParts[2];
                    WebClient client = new WebClient();
                    client.DownloadFile(webAddress, $@"{topLevelImagesFolder}\{sheet}\{itemParts[1]}\{itemParts[0]}.jpg");



                }
                // This else statement will create the sub-folder if it does not exist then add the images to the folder with the product name. 
                else
                {
                    Directory.CreateDirectory(topLevelImagesFolder + "\\" + sheet + "\\" + itemParts[1]);
                    string webAddress = itemParts[2];
                    WebClient client = new WebClient();
                    client.DownloadFile(webAddress, $@"{topLevelImagesFolder}\{sheet}\{itemParts[1]}\{itemParts[0]}.jpg");

                }
            }
            catch (Exception ex)
            {
                if (!triedAlready)
                {
                    triedAlready = true;
                    Thread.Sleep(2000);
                    RetryDownload(topLevelImagesFolder, sheet, itemParts);
                }
                else
                {
                    WriteLine($"An exception has been caught: {ex.Message}. It is on item {itemParts[0]} and the link is {itemParts[2]}.");
                    Task t = ErrorLog(DateTime.Now + "\r" + sheet + "\r" + itemParts[1] + "\r" + itemParts[0] + "\r" + itemParts[2] + "\r");

                }

            }


        }

        public static void RetryDownload(string topLevelImagesFolder, string sheet, List<string> itemParts)
        {
            //Retrying to download the image again.
            DownloadImages(topLevelImagesFolder, sheet, itemParts);

            Task t = ErrorLog(DateTime.Now + " Retried download, check for success" +
                "\r" + sheet + "\r" + itemParts[1] + "\r" + itemParts[0] + "\r" + itemParts[2] + "\r");

            triedAlready = false;

        }


        public static async Task ErrorLog(string errorLogData)
        {
            using StreamWriter file = new(@"C:\Users\jbojovic\Desktop\WebScraperTest\FailedUrls.txt", append: true);
            await file.WriteLineAsync(errorLogData);
        }
    }
}
