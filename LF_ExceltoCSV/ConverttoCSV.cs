using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GroupDocs.Conversion;
using GroupDocs.Conversion.FileTypes;
using GroupDocs.Conversion.Options.Convert;
using GlobalClass;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using SQLTool;
using System.Collections;
using LF_ExceltoCSV.Model;

namespace LF_ExceltoCSV
{
    public class ConverttoCSV
    {
        string ExcelFolder = ConfigurationManager.AppSettings["ExcelFolder"];
        string CSVFolder = ConfigurationManager.AppSettings["CSVFolder"];

        public StringBuilder ConvertExcelToCSV()
        {
            StringBuilder strInfo = new StringBuilder();
            string completeFilePath = string.Empty;
            string err = string.Empty;
            SqlClass.Conn = ConfigurationManager.ConnectionStrings["DataMigrationDB"].ToString();

            try
            {
                string completeFolder = "Complete" + "_" + DateTime.Now.ToString("ddMM");
                

                string getCustomerQry = "select distinct EpicorCustID as CustomerID from tblConvertExcelToSCV";

                IEnumerable<Customer> CustList = SqlClass.SelectQry<Customer>(getCustomerQry, ref err);

                if(CustList != null)
                {
                    foreach(var cust in CustList)
                    {
                        string getCustQry = $"select CustIDFolderName as FolderName, EpicorCustID as CustomerID, ExcelShipToCode as ExcelCode, EpicorShipToCode as EpicorCode, ExcelHeaderCol as ExcelHeader from tblConvertExcelToSCV where EpicorCustID = '{cust.CustomerID.ToString()}'";

                        IEnumerable<Customer> customer = SqlClass.SelectQry<Customer>(getCustQry, ref err);

                        completeFilePath = ExcelFolder + "\\" + customer.FirstOrDefault().FolderName + "\\" + completeFolder;
                        string[] fileList = Directory.GetFiles(ExcelFolder + "\\" + customer.FirstOrDefault().FolderName + "\\");

                        if (fileList.Count() > 0)
                        {
                            foreach (var file in fileList)
                            {
                                ReadExcel(file, customer);

                                using (Converter convert = new Converter(file))
                                {
                                    SpreadsheetConvertOptions options = new SpreadsheetConvertOptions
                                    {
                                        PageNumber = 1,
                                        PagesCount = 1,
                                        Format = SpreadsheetFileType.Csv // Specify the conversion format
                                    };
                                    convert.Convert(CSVFolder + Path.GetFileNameWithoutExtension(file) + ".csv", options);
                                }

                                if (!Directory.Exists(completeFilePath))
                                {
                                    Directory.CreateDirectory(completeFilePath);
                                }

                                //File.Replace(files, completeFilePath + "\\ " + Path.GetFileName(files));
                                MoveWithReplace(file, completeFilePath + "\\ " + Path.GetFileName(file));
                            }
                        }
                        else
                        {
                            //no files
                        }
                    }
                }

                
            }
            catch (Exception ex)
            {
                Log.WriteErrorNlog("Error converting excel to CSV. Message: " + ex.Message.ToString());
            }

            return strInfo;
        }

        public static void MoveWithReplace(string sourceFileName, string destFileName)
        {

            //first, delete target file if exists, as File.Move() does not support overwrite
            if (File.Exists(destFileName))
            {
                File.Delete(destFileName);
            }

            File.Move(sourceFileName, destFileName);

        }

        public void ReadExcel(string filePath, IEnumerable<Customer> customer)
        {
            IWorkbook wb;
            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                string extension = Path.GetExtension(filePath).ToLower();
                if (extension == ".xlsx")
                {
                    wb = new XSSFWorkbook(file);
                }
                else if (extension == ".xls")
                {
                    wb = new HSSFWorkbook(file);
                }
                else
                {
                    string errMsg = "This format is not supported";
                    Log.WriteErrorNlog($"ReadExcel =>StartRead({filePath}) => {errMsg}");
                    throw new Exception(errMsg);
                }
            }

            int ShipToColIdx = 0;
            string shipToCode = string.Empty;

            for (int shIndex = 0; shIndex < wb.NumberOfSheets; shIndex++)
            {
                ISheet sheet = wb.GetSheetAt(shIndex);
                if (sheet.SheetName == "Sheet1")
                {
                    int row = 1;
                    for (int col = 1; col <= sheet.GetRow(row - 1).LastCellNum; col++)
                    {
                        shipToCode = GetCellValue(col, sheet, row);

                        if(shipToCode == "Store Code")
                        {
                            ShipToColIdx = col;
                            break;
                        }
                    }

                    if (ShipToColIdx != 0)
                    {
                        for (row = 2; row <= sheet.LastRowNum; row++)
                        {
                            if(!string.IsNullOrEmpty(GetCellValue(ShipToColIdx, sheet, row)))
                            {
                                shipToCode = GetCellValue(ShipToColIdx, sheet, row);

                                string epiShipToCode = customer.Where(w => w.ExcelCode == shipToCode).Select(s=> s.ExcelCode).ToString();

                                if(!string.IsNullOrEmpty(epiShipToCode))
                                {
                                    //insert new column at the end for epicode
                                    sheet.GetRow(row - 1).CreateCell(sheet.GetRow(row - 1).LastCellNum).SetCellValue(epiShipToCode);
                                }
                            }
                        }
                    }
                }

            }
        }

        private string GetCellValue(int cellIndex, ISheet sheet, int row)
        {
            row = row - 1;
            cellIndex = cellIndex - 1;
            string value = string.Empty;
            string typeStr = string.Empty;
            try
            {
                CellType type = sheet.GetRow(row).GetCell(cellIndex).CellType;
                typeStr = type.ToString();
                switch (typeStr)
                {
                    case "String":
                        value = sheet.GetRow(row).GetCell(cellIndex) == null ? "" : sheet.GetRow(row).GetCell(cellIndex).StringCellValue;
                        break;
                    case "Numeric":
                        value = sheet.GetRow(row).GetCell(cellIndex) == null ? "0" : sheet.GetRow(row).GetCell(cellIndex).NumericCellValue.ToString();
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                Log.WriteErrorNlog($"ReadExcelFile => GetCellValue() => CellIndex:{cellIndex} Row: {row} => ExMsg: {ex.Message} ExStackTrace: {ex.StackTrace}");
                //throw new Exception(ex.Message);
            }

            return value;
        }
    }
}
