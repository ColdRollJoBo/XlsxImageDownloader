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

        static int imagesDownloaded = 0;

        static void Main(string[] args)
        {
            WriteLine("Program Running...");
            // The Spreadsheet Class is the whole excel file at the top most level.
            // The workbook is one step in and is mostly what is used to get actual data.
            Spreadsheet vendorDoc = new Spreadsheet();
            vendorDoc.LoadFromFile(@"C:\Users\jbojovic\Desktop\WebScraperTest\destaco3Column.xlsx");

            Workbook vendorProductCatalog = vendorDoc.Workbook;
            int numberOfSheetsInCatalog = vendorProductCatalog.Worksheets.Count;

            //foreach (string sheet in CollectionOfAllSheetNames(vendorProductCatalog, numberOfSheetsInCatalog))
            //{

            //    LoopDownRowsInCurrentSheetAndDownloadImagesToCorrectFolders(vendorProductCatalog, sheet);

            //}

            // For single Sheet Downloads by sheet name
            // List of Sheets (Manual Clamping, Light-Duty Pneumatic Clamping, Heavy-Duty Pneumatic Clampi - N, NAAMS, Hydraulic Clamping, Indexers, Thrusters Slides, Part Handlers, Conveyors, Rotaries, Grippers - New, Robohand Accessories, End Effectors, Sheet Metal Grippers, Bag Grippers, Tool Changers, Compliance devices)
            LoopDownRowsInCurrentSheetAndDownloadImagesToCorrectFolders(vendorProductCatalog, "Heavy-Duty Pneumatic Clampi - N");

            WriteLine($"{imagesDownloaded} images have been downloaded. There are {vendorProductCatalog.Worksheets.ByName("Heavy-Duty Pneumatic Clampi - N").NotEmptyRowMax} products for download. Please refer to error log for more details.");

        }

        public static List<string> CollectionOfAllSheetNames(Workbook catalog, int sheets)
        {
            List<string> allSheetsInWorkbook = new List<string>();
            // This is a Zero based index so the -1 is needed in the for loop.
            for (int i = 0; i <= sheets - 1; i++)
            {
                string sheet = catalog.Worksheets[i].Name;
                allSheetsInWorkbook.Add(sheet);
                WriteLine(sheet);

            }
            return allSheetsInWorkbook;
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
            for (int i = 1; i <= currentSheet.NotEmptyRowMax; i++)
            {
                for (int j = 0; j <= currentSheet.NotEmptyColumnMax; j++)
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
                    //WriteLine($"{rowInfo[0]} : {rowInfo[1]} : {rowInfo[2]} ");
                    DownloadImages(mainImageFolder, sheet, rowInfo);


                }
                else if (rowInfo.Count < 3)
                {
                    // The +1 is so you get the correct row number in excel because the index is zero based but the row numbers on the side are not. 
                    Task t = ErrorLog($"\rSheet: {sheet} at Row : {i + 1} does not have all 3 parts needed to successfully download an item.");
                    
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
                    string webAddress = itemParts[2].Trim();
                    WebClient client = new WebClient();
                    client.DownloadFile(webAddress, $@"{topLevelImagesFolder}\{sheet}\{itemParts[1]}\{itemParts[0]}.jpg");
                }
                // This else statement will create the sub-folder if it does not exist then add the images to the folder with the product name. 
                else
                {
                    Directory.CreateDirectory(topLevelImagesFolder + "\\" + sheet + "\\" + itemParts[1]);
                    string webAddress = itemParts[2].Trim();
                    WebClient client = new WebClient();
                    client.DownloadFile(webAddress, $@"{topLevelImagesFolder}\{sheet}\{itemParts[1]}\{itemParts[0]}.jpg");

                }
                
                imagesDownloaded++;
            }
            catch (Exception ex)
            {
                WriteLine("Error Caught. Check Error Log for more details.");

                Task t = ErrorLog("\r" + DateTime.Now + "\r" + $"An exception has been caught: {ex.Message}." + "\r" + "Sheet Name: " + sheet + "\r" + "Subfolder: " + itemParts[1] + "\r" + "Item: " + itemParts[0] + "\r" + "URL:" + itemParts[2]);
            }
        }

        public static async Task ErrorLog(string errorLogData)
        {
            using StreamWriter file = new(@"C:\Users\jbojovic\Desktop\WebScraperTest\FailedUrls.txt", append: true);
            await file.WriteLineAsync(errorLogData);
        }
    }
}
