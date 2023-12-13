using TaskToXLSX10._12._23.Data.DTO;

namespace TaskToXLSX10._12._23.Data
{
    public interface IXMLFileWriterCustomer
    {
        /// <summary>
        /// Перезаписывает ячейки у выбранного элемента, строка изменений выберается по Id
        /// </summary>
        /// <param name="indexElement">
        /// Необходимо прибавить +1 для корректного смещения, или +2 если есть шапка документа
        /// </param>
        /// <returns>Bool</returns>
        public bool SaveToXML(int indexElement, Customer customer);

    }
}
