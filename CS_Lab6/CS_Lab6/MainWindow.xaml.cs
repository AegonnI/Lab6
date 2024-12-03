using DocumentFormat.OpenXml.Office2010.Excel;
using System;
using System.IO;
using System.Windows;

namespace CS_Lab6
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ExcelDataBase excelDataBase;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void readDatabase_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                excelDataBase = new ExcelDataBase(Path.GetFullPath(@"..\..\data\LR6-var13.xls"), Path.GetFullPath(@"..\..\data\LR6-var13.xlsx"));
                MessageLabel.Content = "Successful Data Read";
            }
            catch(Exception ex)
            {
                MessageLabel.Content = ex;
            }
            
        }

        private void ShowWorksheet_Click(object sender, RoutedEventArgs e)
        {
            DebugPrint(WorksheetsNames.Items[0].ToString());
            int startIndex = int.Parse(startIndexTextBox.Text);
            int rowsCount = int.Parse(RowsCountTextBox.Text);
            switch (WorksheetsNames.Text)
            {
                case "Счета":
                    DebugPrint(ExcelDataBase.ShowDataBase(excelDataBase.GetAccount(), startIndex, rowsCount));
                    break;
                case "Курс валют":
                    DebugPrint(ExcelDataBase.ShowDataBase(excelDataBase.GetExchangeRate(), startIndex, rowsCount));
                    break;
                case "Поступления":
                    DebugPrint(ExcelDataBase.ShowDataBase(excelDataBase.GetAccrual(), startIndex, rowsCount));
                    break;
            }
        }

        private void DebugPrint(string text)
        {
            MessageLabel.Content = text;
        }

        private void delItemButton_Click(object sender, RoutedEventArgs e)
        {
           int DelID = int.Parse(ID_forDel.Text);
            switch (WorksheetsNames.Text)
            {
                case "Счета":
                    excelDataBase.DelElemInAccounts(DelID);
                    break;
                case "Курс валют":
                    excelDataBase.DelElemInExchangeRates(DelID);
                    break;
                case "Поступления":
                    excelDataBase.DelElemInAccrual(DelID);
                    break;
            }
        }

        private void WorksheetsNames_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (excelDataBase != null)             
            {
                ColoumsNamesComboBox.Items.Clear();
                switch (WorksheetsNames.SelectedItem.ToString().Substring(38))
                {
                    case "Счета":
                        foreach (string item in excelDataBase.GetAccountNames())
                        {
                            ColoumsNamesComboBox.Items.Add(item);
                        }
                        break;
                    case "Курс валют":
                        foreach (string item in excelDataBase.GetExchangeRateNames())
                        {
                            ColoumsNamesComboBox.Items.Add(item);
                        }
                        break;
                    case "Поступления":
                        foreach (string item in excelDataBase.GetAccrualNames())
                        {
                            ColoumsNamesComboBox.Items.Add(item);
                        }
                        break;
                }
            }           
        }

        private void CorrectElemButton_Click(object sender, RoutedEventArgs e)
        {
            int ID = int.Parse(CorrectIDTextBox.Text);
            string column = ColoumsNamesComboBox.Text;
            string rawItem = rawItemTextBox.Text;
            switch (WorksheetsNames.SelectedItem.ToString().Substring(38))
            {
                case "Счета":
                    excelDataBase.CorrectElemInAccount(ID, column, rawItem);
                    break;
                case "Курс валют":
                    excelDataBase.CorrectElemInExchangeRates(ID, column, rawItem);
                    break;
                case "Поступления":
                    excelDataBase.CorrectElemInAccrual(ID, column, rawItem);
                    break;
            }
        }

        private void AddElemButton_Click(object sender, RoutedEventArgs e)
        {
            int ID = int.Parse(AddTextBox1.Text);
            switch (WorksheetsNames.SelectedItem.ToString().Substring(38))
            {
                case "Счета":
                    excelDataBase.AddElemInAccounts(ID, new Account(
                        AddTextBox2.Text, 
                        DateTime.Parse(AddTextBox3.Text)));
                    break;
                case "Курс валют":
                    excelDataBase.AddElemInExchangeRates(ID, new ExchangeRate(
                        AddTextBox2.Text, 
                        double.Parse(AddTextBox3.Text),
                        AddTextBox4.Text));
                    break;
                case "Поступления":
                    excelDataBase.AddElemInAccrual(ID, new Accrual(
                        int.Parse(AddTextBox1.Text),
                        int.Parse(AddTextBox3.Text),
                        DateTime.Parse(AddTextBox4.Text),
                        double.Parse(AddTextBox5.Text)));
                    break;
            }
        }
    }
}
