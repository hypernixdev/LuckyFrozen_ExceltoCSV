using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LF_ExceltoCSV.Model
{
    public class Customer
    {
        public string CustomerID { get; set; } = string.Empty;
        public string FolderName { get; set; } = string.Empty;
        public string ExcelCode { get; set; } = string.Empty;
        public string EpicorCode { get; set; } = string.Empty;
        public string ExcelHeader { get; set; } = string.Empty;

    }
}
