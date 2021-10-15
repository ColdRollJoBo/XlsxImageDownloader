using Bytescout.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using static System.Console;
using ExcelObjects;


namespace XlsxImageDownloader
{
    class Program
    {
        static void Main(string[] args)
        {
            string topLevelImagesFolder = @"C:\Users\jbojovic\Desktop\WebScraperTest\Images";
            XlsxData fileDataParts = new XlsxData();
            fileDataParts.ProductNames = new List<string>();
            fileDataParts.SubProductCategories = new List<string>();
            fileDataParts.Urls = new List<string>();



            WriteLine("Please provide the following information.\n Enter the row number that the headers begin at.");
            // The starting row value has a minus one to account for the zero base indexing that excel uses. 
            fileDataParts.StartingRowForLooping = int.Parse(ReadLine()) - 1;
            WriteLine("Please enter the Column name where the product names are located.");
            fileDataParts.ProductNameColumn = ReadLine().ToLower().Replace(" ", "");
            WriteLine("Please enter the Column name where the sub-folder names are located.");
            fileDataParts.SubProductFolderColumn = ReadLine().ToLower().Replace(" ", "");
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
                WriteLine(fileDataParts.ProductNames.Count);


                LoopDownTheColumnAndAddDataToCorrespondingList(currentSheet, fileDataParts.SubProductFolderColumn,
                   fileDataParts.SubProductFolderColumn, fileDataParts.SubProductCategories, fileDataParts.StartingRowForLooping, lastRowInSheet);
                WriteLine(fileDataParts.SubProductCategories.Count);

                LoopDownTheColumnAndAddDataToCorrespondingList(currentSheet, fileDataParts.UrlColumn,
                   fileDataParts.UrlColumn, fileDataParts.Urls, fileDataParts.StartingRowForLooping, lastRowInSheet);
                WriteLine(fileDataParts.Urls.Count);

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
                HeaderTitle = currentSheet.Cell(loopStartRow, i).ValueAsString.ToLower().Replace(" ", "");

                // TODO Account for headers not being typed exactly (ie: parent category vs parent-category)
                if (HeaderTitle.Equals(listObjectName))
                {

                    // Looping down the columns to add the data to their associated lists.
                    // adding a +1 to pass the Header Row in the sheet.
                    for (int j = loopStartRow + 1; j <= lastRowInSheet; j++)
                    {
                        if (currentSheet.Cell(j, i).ValueAsString != null)
                        {
                            listData.Add(currentSheet.Cell(j, i).ValueAsString);
                        }

                    }
                }
            }

        }



        public static async Task ExampleAsync(string errorLogData)
        {
            using StreamWriter file = new(@"C:\Users\jbojovic\Desktop\FailedUrls.txt", append: true);
            await file.WriteLineAsync(errorLogData);
        }






    }
}
