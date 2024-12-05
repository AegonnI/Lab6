# Лабораторная работа №6. LINQ-запросы. Вариант 13
Разработать консольное приложение с дружественным интерфейсом с возможностью выбора
заданий. Приложение должно выполнять следующие функции:
1. Чтение базы данных из excel файла.
2. Просмотр базы данных.
3. Удаление элементов (по ключу).
4. Корректировка элементов (по ключу).
5. Добавление элементов.
6. Реализация 4 запросов (формулировки запросов придумать самостоятельно и отразить в
отчёте, можно использовать запрос, данный в примере):  
  1 запрос с обращением к одной таблице(Запрос1 - Сколько было начислений у каждой валюты(ID))  
  1 запрос с обращением к двум таблицам(Запрос2 - Общий доход у каждой валюты)
  2 запроса с обращением к трем таблицам(Запрос3 - Держатель с минимальными начислениями(включая отрицательные) и само начисление,
   (Запрос4 - Держатель с максимальными начислениями и само начисление)
2 запроса должны возвращать перечень, 2 запроса одно значение.
8. Во время всего сеанса работы ведется полное протоколирование действий в текстовом
файле (в начале сеанса запросить, будет ли это новый файл или дописывать в уже
существующий).   

Элементами базы данных являются объекты классов согласно вашему варианту. Содержание классов
определить самостоятельно и отразить в отчете (в классах должны присутствовать свойства,
конструкторы, перегруженный метод ToString). Весь функционал приложения реализовать в виде
методов вспомогательного класса с помощью LINQ-запросов.
Предусмотреть обработку возможных ошибок при работе программы.

В файле LR6-var13.xls приведён фрагмент базы данных «Инвестиционные счета». Таблица «Счета»
содержит информацию о владельце счёта и дате его открытия. Таблица «Курс валют» содержит
информацию о курсах валют по отношению к рублю. Таблица «Начисления» содержит информацию о
всех операциях со счетом: код счёта, код валюты, дату операции и сумму начисления (она может быть
отрицательной).

## Классы
## Класс Account
Реализует первый лист эксель файла

## Поля
```c#
public string fullName;
public DateTime date;
```

## Конструторы
## Конструтор по умолчанию
```c#
public Account()
{
    fullName = "";
    date = DateTime.MinValue;
}
```
## Конструтор присваивания
```c#
public Account(string fullName, DateTime date)
{
    this.fullName = fullName;
    this.date = date;
}
```

## Метод
```c#
public override string ToString()
```

## Класс ExchangeRate
Реализует второй лист эксель файла

## Поля
```c#
public string letterCode;
public double exchangeRate;
public string fullName;
```

## Конструторы
## Конструтор по умолчанию
```c#
public ExchangeRate()
{
    letterCode = "";
    exchangeRate = 0;
    fullName = "";
}
```
## Конструтор присваивания
```c#
public ExchangeRate(string letterCode, double exchangeRate, string fullName)
{
    this.letterCode = letterCode;
    this.exchangeRate = exchangeRate;
    this.fullName = fullName;
}
```

## Метод
```c#
public override string ToString()
```

## Класс Accrual
Реализует третий лист эксель файла

## Поля
```c#
public int accountID;
public int currencyID;
public DateTime date;
public double summ;
```

## Конструторы
## Конструтор по умолчанию
```c#
public Accrual()
{
    accountID = 0;
    currencyID = 0;
    date = DateTime.MinValue;
    summ = 0;
}
```
## Конструтор присваивания
```c#
public Accrual(int accountID, int currencyID, DateTime date, double summ)
{
    this.accountID = accountID;
    this.currencyID = currencyID;
    this.date = date;
    this.summ = summ;
}
```

## Метод
```c#
public override string ToString()
```

## Класс ExcelDataBase
Основной класс взаимодействия

## Поля
```c#
public Dictionary<int, Account> accounts { get; }
public Dictionary<int, ExchangeRate> exchangeRates { get; }
public Dictionary<int, Accrual> accruals { get; }

public List<string> accountNames { get; }
public List<string> exchangeRateNames { get; }
public List<string> accrualNames { get; }
```

## Конструторы
## Конструтор по умолчанию
```c#
public ExcelDataBase()
{
    accounts = new Dictionary<int, Account>();
    exchangeRateNames = new List<string>();
    accruals = new Dictionary<int, Accrual>();

    accrualNames = new List<string>();
    exchangeRateNames = new List<string>();
    accrualNames = new List<string>();
}
```
## Основной конструтор
```c#
public ExcelDataBase(string pathXLS, string pathXLSX)
{
    if (!File.Exists(pathXLS)) throw new Exception();

    if (!File.Exists(pathXLSX))
    {
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(pathXLS);
        workbook.SaveAs(pathXLSX, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
        workbook.Close();
        excelApp.Quit();
    }

    try
    {
        using (XLWorkbook wb = new XLWorkbook(pathXLSX))
        {

            IXLWorksheet ws = wb.Worksheet(1);

            accountNames = ws.Row(1).CellsUsed().Select(cell => cell.GetText()).ToList();

            accounts = ws.RowsUsed()
                                  .Skip(1)
                                  .ToDictionary(
                                      row => (int)row.Cell(1).Value.GetNumber(),
                                      row => new Account(row.Cell(2).GetText(), row.Cell(3).GetDateTime())
                                  );

            ws = wb.Worksheet(2);

            exchangeRateNames = ws.Row(1).CellsUsed().Select(cell => cell.GetText()).ToList();

            exchangeRates = ws.RowsUsed()
                                  .Skip(1)
                                  .ToDictionary(
                                      row => (int)row.Cell(1).Value.GetNumber(),
                                      row => new ExchangeRate(row.Cell(2).GetText(), row.Cell(3).Value.GetNumber(), row.Cell(4).GetText().Trim())
                                  );

            ws = wb.Worksheet(3);

            accrualNames = ws.Row(1).CellsUsed().Select(cell => cell.GetText()).ToList();

            accruals = ws.RowsUsed()
                                  .Skip(1)
                                  .ToDictionary(
                                      row => (int)row.Cell(1).Value.GetNumber(),
                                      row => new Accrual((int)row.Cell(2).Value.GetNumber(), (int)row.Cell(3).Value.GetNumber(), row.Cell(4).GetDateTime(), row.Cell(5).Value.GetNumber())
                                  );
        }
    }
    catch (Exception ex)
    {
        throw ex;
    }
}
```

## Методы
```c#
public void Save(string pathXLSX)

//Запросы
public string HighestAccruedAccountHolder()
public string TheMostLostAccountHolder()
public string IncomeCurrencies()
public string CountAccrualsForAllCur()

//Удаление
public void DelElemInAccounts(int ID)
public void DelElemInExchangeRates(int ID)
public void DelElemInAccrual(int ID)

//Корректировка
public void CorrectElemInAccount(int ID, string column, string rawItem)
public void CorrectElemInExchangeRates(int ID, string column, string rawItem)
public void CorrectElemInAccrual(int ID, string column, string rawItem)

//Добавление
public void AddElemInAccounts(int ID, Account account)
public void AddElemInExchangeRates(int ID, ExchangeRate exchangeRate)
public void AddElemInAccrual(int ID, Accrual accrual)

//ToString
public static string ToString<T>(Dictionary<int, T> dict, int startIndex,int rowsCount) where T : class
```

## Тесты
# Пользователь выбриает "В новом файле"
![image](https://github.com/user-attachments/assets/90bead8b-c356-4107-bb92-9fe470f2b499)
# В случае неудачи выводим сообщение
![image](https://github.com/user-attachments/assets/9daa3da4-b616-49e4-bda4-7fda6e6c7a32)
# В случае удачного чтения, открываем пользователю весь функционал
![image](https://github.com/user-attachments/assets/6aa4df6a-138a-485f-ac9d-49ceece38e82)
# Просмотр базы данных
![image](https://github.com/user-attachments/assets/cf5a1bac-2e68-43b7-946c-1f164b208a69)
# Удаление элемента с ключем 5
![image](https://github.com/user-attachments/assets/116487fe-3b27-4dfe-97d4-bf4a83f363fd)
# Корректировка элемента с ключем 8, замена столбца ФИО на Понд
![image](https://github.com/user-attachments/assets/6788c4a8-52bc-41ce-81bc-e4fa737b5c8b)\
# Добавление элемента
![image](https://github.com/user-attachments/assets/6712d8ed-711a-4de2-8edb-a730366880e5)

# Тест запроса 1
![image](https://github.com/user-attachments/assets/7fe8af5d-56be-4fda-bf12-3da5221f12f9)
# Тест запроса 2 
![image](https://github.com/user-attachments/assets/4b5ec9fb-d71b-4257-8f9a-95400a55c0cb)
# Тест запроса 3
![image](https://github.com/user-attachments/assets/353be4bf-4fa7-470f-9529-8607d77242a7)
# Тест запроса 4
![image](https://github.com/user-attachments/assets/cbaa5fdc-3e77-402b-b3b7-9da81b029d29)

# Пример протокола
![image](https://github.com/user-attachments/assets/b7315da1-ab19-4551-8934-46d591795f82)
