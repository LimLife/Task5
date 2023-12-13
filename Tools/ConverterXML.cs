using ClosedXML.Excel;
using TaskToXLSX10._12._23.Data.DTO;

namespace TaskToXLSX10._12._23.Tools
{
    public static class ConverterXML
    {
        /// <summary>
        /// Конвертация строки о ключам к объекту по типу
        /// </summary>
        /// <param name="type">Тип ковертируемого объекта</param>
        /// <param name="value">Объект строки ячеек</param>
        /// <returns>Объект</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object ConvertByType(Type type, IXLRangeRow value)
        {
            if (type == typeof(Product))
                return ToProduct(value);
            if (type == typeof(Order))
                return ToOrder(value);
            if (type == typeof(Customer))
                return ToCustomer(value);
            else
                throw new ArgumentException("Не известный тип");
        }

        public static Product ToProduct(IXLRangeRow row)
        {
            return new Product
            {
                Id = Convert.ToInt32(row.Cell((int)ProductColumns.Id).Value.GetNumber()),
                Name = row.Cell((int)ProductColumns.Name).Value.ToString(),
                Unist = row.Cell((int)ProductColumns.Unist).Value.ToString(),
                PricePerPiece = Convert.ToInt32(row.Cell((int)ProductColumns.PricePerPiece).Value.GetNumber())
            };
        }
        public static Order ToOrder(IXLRangeRow row)
        {
            return new Order
            {
                Id = Convert.ToInt32(row.Cell((int)OrderColumns.Id).Value.GetNumber()),
                IdProduc = Convert.ToInt32(row.Cell((int)OrderColumns.IdProduc).Value.GetNumber()),
                IdCustomer = Convert.ToInt32(row.Cell((int)OrderColumns.IdCustomer).Value.GetNumber()),
                IdOrder = Convert.ToInt32(row.Cell((int)OrderColumns.IdOrder).Value.GetNumber()),
                RequiredQuantity = Convert.ToInt32(row.Cell((int)OrderColumns.RequiredQuantity).Value.GetNumber()),
                PublicationDate = row.Cell((int)OrderColumns.PublicationDate).Value.GetDateTime(),
            };
        }
        public static Customer ToCustomer(IXLRangeRow row)
        {
            return new Customer
            {
                Id = Convert.ToInt32(row.Cell((int)CustomerColumns.Id).Value.GetNumber()),
                NameCompany = row.Cell((int)CustomerColumns.NameCompany).Value.ToString(),
                Address = row.Cell((int)CustomerColumns.Address).Value.ToString(),
                Manager = row.Cell((int)CustomerColumns.Manager).Value.ToString()
            };
        }
    }
}
