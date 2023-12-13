using TaskToXLSX10._12._23.Data.DTO;
using TaskToXLSX10._12._23.Tools;
using ClosedXML.Excel;

namespace TaskToXLSX10._12._23.Data
{
    public class FileEditor : IFileEditor
    {
        private string _filePath = "";
        public List<T> GetDataFromTable<T>(int numberSheetList) where T : class
        {
            using (var excelWorkbook = new XLWorkbook(_filePath))
            {
                try
                {
                    var worksheet = excelWorkbook.Worksheet(numberSheetList);
                    var firstCell = worksheet.FirstCellUsed();
                    var lastCell = worksheet.LastCellUsed();
                    var range = worksheet.Range(firstCell.Address, lastCell.Address);
                    var type = typeof(T);
                    //Пропуск шапки документа, где находятся описания колонок.
                    var result = range.Rows().Skip(1).Select(item => (T)ConverterXML.ConvertByType(type, item)).ToList();
                    return (List<T>)result;

                }
                catch (IOException ex)
                {
                    Console.WriteLine("Файл не доступен");
                    return [];
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    return [];
                }
            }
        }
        public bool SaveToXML(int row, Customer customer)
        {
            using var workBook = new XLWorkbook(_filePath);
            try
            {
                var worksheet = workBook.Worksheet(ExcelSheetConstants.CustomerWorksheet);
                var rowCells = worksheet.Row(row);
                rowCells.Cell((int)CustomerColumns.NameCompany).Value = (XLCellValue)customer.NameCompany;
                rowCells.Cell((int)CustomerColumns.Manager).Value = (XLCellValue)customer.Manager;
                workBook.Save();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        public void SetPathToFileString(string path)
        {
            _filePath = path;
        }

    }
}
