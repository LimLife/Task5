using System.Text;
using TaskToXLSX10._12._23.Data;
using TaskToXLSX10._12._23.Data.DTO;
using System.Text.RegularExpressions;
using TaskToXLSX10._12._23.Data.Enums;


namespace TaskToXLSX10._12._23.Application
{
    public class Application : IApplicationInput
    {
        public required IFileEditor _editor { get; init; }

        private ChooseExecute _state = ChooseExecute.Choose;
        public void Input()
        {
            SetPathToFile();
            while (true)
            {
                switch (_state)
                {
                    case ChooseExecute.Choose:
                        Console.WriteLine("Введите нужный режим.");
                        ChooseMenu();
                        string chooce = Console.ReadLine();
                        var seccess = Enum.TryParse(chooce, out ChooseExecute execute);
                        if (!seccess)
                            return;
                        _state = execute;
                        break;
                    case ChooseExecute.ChangePerson:
                        ChangeCustomerContactPerson();
                        _state = ChooseExecute.Choose;
                        break;
                    case ChooseExecute.FindGolden:
                        FindGoldenCustomer();
                        _state = ChooseExecute.Choose;
                        break;
                    case ChooseExecute.GetInformation:
                        GetInformationAboutProduct();
                        _state = ChooseExecute.Choose;
                        break;
                    case ChooseExecute.Exit:
                        Environment.Exit(0);
                        break;
                }
            }
        }

        private void ChangeCustomerContactPerson()
        {
            var customers = _editor.GetDataFromTable<Customer>(ExcelSheetConstants.CustomerWorksheet);
            var builder = new StringBuilder();
            builder.AppendLine("| Код клиента | Наименование организации | Адрес | Контактное лицо (ФИО) |");
            builder.AppendLine("|-------------|--------------------------|-------|-----------------------|");
            foreach (var client in customers)
            {
                builder.AppendLine($"| {client.Id} | {client.NameCompany} | {client.Address} | {client.Manager}|");
            }
            Console.WriteLine(builder.ToString());
            Console.WriteLine("Введите название организации которую хотите изменить.");
            string companyName = Console.ReadLine();
            if (string.IsNullOrEmpty(companyName))
            {
                Console.WriteLine("Пустое значение!!!");
                return;
            }

            var customer = customers.Where(x => x.NameCompany == companyName).FirstOrDefault();
            if (customer is null)
            {
                Console.WriteLine("Такой организации не существует");
                return;
            }
            Console.WriteLine("Введите новое название организации");
            var newCompanyName = Console.ReadLine();
            if (string.IsNullOrEmpty(newCompanyName))
            {
                Console.WriteLine("Пустое значение!!!");
                return;
            }
            Console.WriteLine("Введите новое контактное лицо введите его ФИО");
            var contactPersonName = Console.ReadLine();
            if (string.IsNullOrEmpty(contactPersonName)) return;
            string pattern = "^[А-ЯЁа-яё]+[-' ]?[А-ЯЁа-яё]+[-' ]?[А-ЯЁа-яё]+$";
            Match match = Regex.Match(contactPersonName, pattern);
            if (match.Success)
            {
                customer.NameCompany = newCompanyName;
                customer.Manager = contactPersonName;

                // +2 Для правильного остчента 0 элемент +1, а также с учетом шапки +1  
                if (_editor.SaveToXML(customers.IndexOf(customer) + 2, customer))
                {
                    Console.WriteLine($"Были изменено поле Наименование организации на : {customer.NameCompany} ");
                    Console.WriteLine($"Были изменено поле Контактное лицо (ФИО)    на : {customer.Manager}");
                }
            }
            else
                Console.WriteLine("Некорректное ФИО");

        }
        private void FindGoldenCustomer()
        {
            var orders = _editor.GetDataFromTable<Order>(ExcelSheetConstants.OrderWorksheet);

            var goldYear = (from order in orders orderby order.IdCustomer select order).First();
            var goldMonth = (from order in orders orderby order.IdCustomer, order.PublicationDate select order).OrderBy(data => data.PublicationDate).First();

            var customer = _editor.GetDataFromTable<Customer>(ExcelSheetConstants.CustomerWorksheet);

            var clinetYear = customer.FirstOrDefault(client => client.Id == goldYear.IdCustomer);
            var clientMonth = customer.FirstOrDefault(client => client.Id == goldMonth.IdCustomer);

            {
                Console.WriteLine($"Золотой клиент за год: {clinetYear?.NameCompany ?? "Не найден"}");
                Console.WriteLine($"Золотой клиент за месяц: {clientMonth?.NameCompany ?? "Не найден"}");
            }
        }
        private void GetInformationAboutProduct()
        {
            var products = _editor.GetDataFromTable<Product>(ExcelSheetConstants.ProductWorksheet);
            var builder = new StringBuilder();
            builder.AppendLine("| Код товара | Наименование | Ед. измерения | Цена товара за единицу |");
            builder.AppendLine("|------------|--------------|---------------|------------------------|");
            foreach (var product in products)
            {
                builder.AppendLine($"| {product.Id} | {product.Name} | {product.Unist} | {product.PricePerPiece}|");
            }
            Console.WriteLine(builder.ToString());
            Console.WriteLine("Введите название товара");
            var nameProduct = Console.ReadLine();

            var productChoose = products.Find(produtName => produtName.Name == nameProduct);
            if (productChoose is null)
            {
                Console.WriteLine("Товара нет в списке");
                return;
            }
            var orders = _editor.GetDataFromTable<Order>(ExcelSheetConstants.OrderWorksheet);
            var customers = _editor.GetDataFromTable<Customer>(ExcelSheetConstants.CustomerWorksheet);
            var filter = (from order in orders
                          join client in customers on order.IdCustomer equals client.Id
                          select new
                          {
                              Id = order.Id,
                              ProductName = productChoose.Name,
                              PricePerPiece = productChoose.PricePerPiece,
                              PublicationDate = order.PublicationDate,
                              RequiredQuantity = order.RequiredQuantity,
                              NameCompany = client.NameCompany,
                              Address = client.Address,
                              Manager = client.Manager
                          }).DistinctBy(x => x.NameCompany);

            builder.Clear();
            builder.AppendLine("|Код клента |Наименования организации|   Адресс   |  Контактное лицо      |   Название товара   |Цена за еденицу|  Дата заказа   |Колличечество товаров");
            builder.AppendLine("|-----------|------------------------|------------|-----------------------|---------------------|---------------|----------------|----------------------|");
            foreach (var dataFilter in filter)
            {
                builder.AppendLine($"|{dataFilter.Id}| {dataFilter.NameCompany} | {dataFilter.Address} | {dataFilter.Manager} | {dataFilter.ProductName} | {dataFilter.PricePerPiece} | {dataFilter.PublicationDate} | {dataFilter.RequiredQuantity} ");
            }
            Console.WriteLine(builder.ToString());
        }

        private void SetPathToFile()
        {
            Console.WriteLine("Введите абсолютный путь до файла пример: D:\\\\Folder\\\\your.xlsx");
            var @path = Console.ReadLine();
            if (IsExistFile(path) is false)
                SetPathToFile();
            _editor.SetPathToFileString(path);
        }
        private bool IsExistFile(string path)
        {
            if (File.Exists(path) && Path.GetExtension(path) == ".xlsx")
            {
                return true;
            }
            Console.WriteLine("Упс...  что-то не так документ не найден попробуйте еще раз.");
            return false;
        }
        private void ChooseMenu()
        {
            Console.WriteLine("\t 1 Для редактирования данных компании `Названия организации и котактного лица.` \n " +
                "\t 2 Для вывод залотого клиента за год и месяц.\n \t 3.Для получение информации по названиб товара\n \t 4.Для выхода из программы ");
        }

    }
}
