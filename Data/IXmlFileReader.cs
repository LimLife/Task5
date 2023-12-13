using ClosedXML.Excel;

namespace TaskToXLSX10._12._23.Data
{
    public interface IXmlFileReader
    {
        /// <summary>
        /// Метод возвращает список значений из указанной таблицы. 
        /// <param name="numberSheetList">
        /// Таблиа из которой необходимо получить значения.
        /// </param>
        /// <returns>
        /// Список значений из указанной таблицы.
        /// </returns>
        /// </summary>
        public List<T> GetDataFromTable<T>(int numberSheetList) where T : class;
    }
}
