using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CS_Lab6
{
    internal class ExcelDataBase
    {
        Dictionary<int, Account> accounts { get; }
        Dictionary<int, ExchangeRate> exchangeRates { get; }
        Dictionary<int, Accrual> accruals { get; }

        List<string> accountNames { get; }
        List<string> exchangeRateNames { get; }
        List<string> accrualNames { get; }

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

        public string HighestAccruedAccountHolder()
        {
            Dictionary<int, double> HolderAnSumm = accruals
                        .GroupBy(item => item.Value.accountID)
                        .ToDictionary(
                            group => group.Key,
                            group => group.Sum(item => item.Value.summ * exchangeRates[item.Value.currencyID].exchangeRate));          
            
            return new string(accounts[HolderAnSumm
                        .OrderByDescending(x => x.Value)
                        .First().Key].fullName
                            .TakeWhile(s => s != ' ')
                            .ToArray())
                            .ToUpper() + " " +
                         HolderAnSumm
                            .OrderByDescending(x => x.Value)
                            .First().Value;
        }

        public string TheMostLostAccountHolder()
        {
            Dictionary<int, double> HolderAnSumm = accruals
                        .GroupBy(item => item.Value.accountID)
                        .ToDictionary(
                            group => group.Key,
                            group => group.Sum(item => item.Value.summ * exchangeRates[item.Value.currencyID].exchangeRate));

            return new string(accounts[HolderAnSumm
                        .OrderBy(x => x.Value)
                        .First().Key].fullName
                            .TakeWhile(s => s != ' ')
                            .ToArray())
                            .ToUpper() 
                            + " " +
                         HolderAnSumm
                            .OrderBy(x => x.Value)
                            .First().Value;
        }

        public string IncomeCurrencies()
        {
            return string.Join("\n", accruals
                .GroupBy(item => item.Value.currencyID)
                .ToDictionary(
                    group => group.Key,
                    group => group.Sum(item => item.Value.summ * exchangeRates[item.Value.currencyID].exchangeRate)               )
                .Select(x => exchangeRates[x.Key].letterCode + " " + x.Value.ToString())
                .ToList());
        }

        public string CountAccrualsForAllCur()
        {
            return string.Join("\n", accruals
                .GroupBy(item => item.Value.currencyID)
                .ToDictionary(
                    group => group.Key,
                    group => group.Sum(item => 1))
                .Select(x => x.Key.ToString() + " " + x.Value.ToString())
                .ToList());
        }

        public static string ShowDataBase<T>(Dictionary<int, T> dict, int startIndex,int rowsCount) where T : class
        {
            string result = "";
            int i = 0;
            int j = 0;

            foreach (int id in dict.Keys)
            {
                if (i == rowsCount) break;

                if (j == startIndex)
                {
                    i++;
                    result += id.ToString() + " " + dict[id].ToString() + "\n";
                }
                else j++;

            }
            result += "...\n";

            return result + $"Всего {dict.Count} строк";
        }

        public void DelElemInAccounts(int key)
        {
            try { accounts.Remove(key); }
            catch (Exception ex) { throw ex; }
        }

        public void DelElemInExchangeRates(int key)
        {
            try { exchangeRates.Remove(key); }
            catch (Exception ex) { throw ex; }
        }

        public void DelElemInAccrual(int key)
        {
            try { accruals.Remove(key); }
            catch (Exception ex) { throw ex; }
        }

        public void CorrectElemInAccount(int ID, string column, string rawItem) 
        {
            switch (column)
            {
                case "ID":
                    accounts[int.Parse(rawItem)] = accounts[ID];
                    accounts.Remove(ID);
                    break;
                case "ФИО":
                    accounts[ID].fullName = rawItem;
                    break;
                case "Дата открытия вклада":
                    accounts[ID].date = DateTime.Parse(rawItem);
                    break;
            }
        }

        public void CorrectElemInExchangeRates(int ID, string column, string rawItem)
        {
            switch (column)
            {
                case "ID":
                    exchangeRates[int.Parse(rawItem)] = exchangeRates[ID];
                    exchangeRates.Remove(ID);
                    break;
                case "Буквенный код":
                    exchangeRates[ID].letterCode = rawItem;
                    break;
                case "Курс":
                    exchangeRates[ID].exchangeRate = double.Parse(rawItem);
                    break;
                case "Полное наименование":
                    exchangeRates[ID].fullName = rawItem;
                    break;
            }
        }

        public void CorrectElemInAccrual(int ID, string column, string rawItem)
        {
            switch (column)
            {
                case "ID":
                    accruals[int.Parse(rawItem)] = accruals[ID];
                    accruals.Remove(ID);
                    break;
                case "ID счёта":
                    accruals[ID].accountID = int.Parse(rawItem);
                    break;
                case "ID валюты":
                    accruals[ID].currencyID = int.Parse(rawItem);
                    break;
                case "Дата":
                    accruals[ID].date = DateTime.Parse(rawItem);
                    break;
                case "Сумма":
                    accruals[ID].summ = double.Parse(rawItem);
                    break;
            }
        }

        public void AddElemInAccounts(int ID, Account account) { accounts[ID] = account; }

        public void AddElemInExchangeRates(int ID, ExchangeRate exchangeRate) { exchangeRates[ID] = exchangeRate; }

        public void AddElemInAccrual(int ID, Accrual accrual) { accruals[ID] = accrual; }

        public Dictionary<int, Account> GetAccount() { return accounts; }
        public Dictionary<int, ExchangeRate> GetExchangeRate() { return exchangeRates; }
        public Dictionary<int, Accrual> GetAccrual() { return accruals; }

        public List<string> GetAccountNames() { return accountNames; }
        public List<string> GetExchangeRateNames() { return exchangeRateNames; }
        public List<string> GetAccrualNames() { return accrualNames; }
    }
}
