using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlsxImageDownloader
{
    class SheetDataSections
    {
        public List<string> MainProductCategories { get; set; }
        public List<string> SubProductCategories { get; set; }
        public List<string> ProductNames { get; set; }
        public List<string> Urls { get; set; }
        public string UrlColumn { get; set; }
        public string ProductNameColumn { get; set; }
        public string SubProductFolderColumn { get; set; }
        public int StartingRowForLooping { get; set; }

        public SheetDataSections()
        {

        }

    }
}
