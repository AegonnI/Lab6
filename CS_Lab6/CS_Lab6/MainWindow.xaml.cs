using DocumentFormat.OpenXml.Office2010.Excel;
using System;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace CS_Lab6
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ExcelDataBase excelDataBase;

        public string protocol;

        public MainWindow()
        {
            InitializeComponent();

            protocol = "";

            referenceLabel.Content = "Справка:" +
                "\nЗапрос1 - Сколько было \nначислений у каждой валюты(ID)" +
                "\nЗапрос2 - Общий доход \nу каждой валюты" +
                "\nЗапрос3 - Держатель с \nминимальными начислениями\n(включая отрицательные)" +
                "\nЗапрос4 - Держатель с \nмаксимальными начислениями";

            ChangeVisible(false, ShowWorksheet, delItemButton, CorrectElemButton, AddElemButton, RequestButton1, RequestButton2, RequestButton3, RequestButton4, Reset);
            ChangeVisible(false, startIndexTextBox, RowsCountTextBox, ID_forDel, CorrectIDTextBox, rawItemTextBox, AddTextBox1, AddTextBox2, AddTextBox3, AddTextBox4, AddTextBox5);
            ChangeVisible(false, ColoumsNamesComboBox, WorksheetsNames);
            ChangeVisible(false, OutputLabel, LeftLabel1, LeftLabel2, LeftLabel3, LeftLabel4, LeftLabel5, LeftLabel6, LeftLabel7, LeftLabel8, LeftLabel9, LeftLabel10, LeftLabel11, LeftLabel12, referenceLabel);
        }

        void DataWindow_Closing(object sender, CancelEventArgs e)
        {
            AddToProtocol("Clicked Close Window");
            AddToProtocol("Window is closed");
            File.WriteAllText(Path.GetFullPath(@"..\..\data\Protocol.txt"), protocol);
        }

        private void ChangeVisible<T>(bool enable, params T[] values) where T : System.Windows.Controls.Control
        {
            for (int i = 0; i < values.Length; i++)
            {
                values[i].Visibility = enable ? Visibility.Visible : Visibility.Hidden;
            }
        }

        private void EnableControls<T>(bool enable, params T[] values) where T : System.Windows.Controls.Control
        {
            for (int i = 0; i < values.Length; i++) 
            {
                values[i].IsEnabled = enable;
            }
        }

        private void Switcher(string condition, Action action1, Action action2, Action action3)
        {
            switch (condition)
            {
                case "Счета":
                    action1();
                    break;
                case "Курс валют":
                    action2();
                    break;
                case "Поступления":
                    action3();
                    break;
            }
        }

        private void PrintWorksheet(int startIndex, int rowsCount)
        {
            Switcher(WorksheetsNames.Text,
                () => DebugPrint(ExcelDataBase.ShowDataBase(excelDataBase.GetAccount(), startIndex, rowsCount)),
                () => DebugPrint(ExcelDataBase.ShowDataBase(excelDataBase.GetExchangeRate(), startIndex, rowsCount)),
                () => DebugPrint(ExcelDataBase.ShowDataBase(excelDataBase.GetAccrual(), startIndex, rowsCount)));
        }

        

        private void ShowWorksheet_Click(object sender, RoutedEventArgs e)
        {
            AddToProtocol("ShowWorksheet was Clicked");

            //try { int.Parse(startIndexTextBox.Text); }
            //catch 
            //{
            //    AddToProtocol("ShowWorksheet was Clicked");
            //}

            try
            {
                PrintWorksheet(int.Parse(startIndexTextBox.Text), int.Parse(RowsCountTextBox.Text));
                AddToProtocol("Worksheet was written");
            }
            catch
            {
                OutputTextBox.Text = "Cant Print Worksheet";
                AddToProtocol("Cant Print Worksheet");
                return;
            }           
        }

        private void DebugPrint(string text)
        {
            OutputTextBox.Text = text;
        }

        private void delItemButton_Click(object sender, RoutedEventArgs e)
        {
            AddToProtocol("delItemButton was clicked");

            try
            {
                Switcher(WorksheetsNames.Text,
                    () => excelDataBase.DelElemInAccounts(int.Parse(ID_forDel.Text)),
                    () => excelDataBase.DelElemInExchangeRates(int.Parse(ID_forDel.Text)),
                    () => excelDataBase.DelElemInAccrual(int.Parse(ID_forDel.Text)));

                AddToProtocol("Seccessful delete element");
            }
            catch
            {
                OutputTextBox.Text = "Cant delete element";
                AddToProtocol("Cant delete element");
                return;
            }

            try
            {
                PrintWorksheet(int.Parse(ID_forDel.Text) - int.Parse(RowsCountTextBox.Text) / 2 >= 0 ?
                    int.Parse(ID_forDel.Text) - int.Parse(RowsCountTextBox.Text) / 2 : 0,
                    int.Parse(RowsCountTextBox.Text));
                AddToProtocol("Worksheet was written");
            }
            catch
            {
                OutputTextBox.Text = "Cant Print Worksheet";
                AddToProtocol("Cant Print Worksheet");
                return;
            }


        }

        private void WorksheetsNames_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            //AddToProtocol("WorksheetsNames was clicked");

            try
            {
                if (excelDataBase != null)
                {
                    ColoumsNamesComboBox.Items.Clear();
                    Switcher(WorksheetsNames.SelectedItem.ToString().Substring(38),
                        () => {
                            foreach (string item in excelDataBase.GetAccountNames())
                                ColoumsNamesComboBox.Items.Add(item);
                            ChangeVisible(false, AddTextBox4, AddTextBox5);
                            LeftLabel8.Content = "ID";
                            LeftLabel9.Content = "ФИО";
                            LeftLabel10.Content = "Дата открытия вклада";
                        },
                        () => {
                            foreach (string item in excelDataBase.GetExchangeRateNames())
                                ColoumsNamesComboBox.Items.Add(item);
                            ChangeVisible(true, AddTextBox4);
                            ChangeVisible(false, AddTextBox5);
                            LeftLabel8.Content = "ID";
                            LeftLabel9.Content = "Буквенный код";
                            LeftLabel10.Content = "Курс";
                            LeftLabel11.Content = "Полное наименование";
                        },
                        () => {
                            foreach (string item in excelDataBase.GetAccrualNames())
                                ColoumsNamesComboBox.Items.Add(item);
                            ChangeVisible(true, AddTextBox4, AddTextBox5);
                            LeftLabel8.Content = "ID";
                            LeftLabel9.Content = "ID счёта";
                            LeftLabel10.Content = "ID валюты";
                            LeftLabel11.Content = "Дата";
                            LeftLabel12.Content = "Сумма";
                        });
                }

                //AddToProtocol("WorksheetsNames was clicked");
            }
            catch
            {
                OutputTextBox.Text = "Cant Create WorksheetsNames items";
                AddToProtocol("Cant Create WorksheetsNames items");
                return;
            }

        }

        private void CorrectElemButton_Click(object sender, RoutedEventArgs e)
        {
            AddToProtocol("CorrectElemButton was clicked");

            try
            {
                int ID = int.Parse(CorrectIDTextBox.Text);
                string column = ColoumsNamesComboBox.Text;
                string rawItem = rawItemTextBox.Text;

                Switcher(WorksheetsNames.SelectedItem.ToString().Substring(38),
                        () => excelDataBase.CorrectElemInAccount(ID, column, rawItem),
                        () => excelDataBase.CorrectElemInExchangeRates(ID, column, rawItem),
                        () => excelDataBase.CorrectElemInAccrual(ID, column, rawItem));

                PrintWorksheet(ID - int.Parse(RowsCountTextBox.Text) / 2 >= 0 ?
                    ID - int.Parse(RowsCountTextBox.Text) / 2 : 0, int.Parse(RowsCountTextBox.Text));

                AddToProtocol("Seccessful Correct");
            }
            catch
            {
                OutputTextBox.Text = "Cant Print Worksheet";
                AddToProtocol("WorksheetsNames was clicked");
                return;
            }



        }

        private void AddElemButton_Click(object sender, RoutedEventArgs e)
        {
            int ID = int.Parse(AddTextBox1.Text);

            Switcher(WorksheetsNames.SelectedItem.ToString().Substring(38),
                    () => excelDataBase.AddElemInAccounts(ID, new Account(
                        AddTextBox2.Text,
                        DateTime.Parse(AddTextBox3.Text))),
                    () => excelDataBase.AddElemInExchangeRates(ID, new ExchangeRate(
                        AddTextBox2.Text,
                        double.Parse(AddTextBox3.Text),
                        AddTextBox4.Text)),
                    () => excelDataBase.AddElemInAccrual(ID, new Accrual(
                        int.Parse(AddTextBox1.Text),
                        int.Parse(AddTextBox3.Text),
                        DateTime.Parse(AddTextBox4.Text),
                        double.Parse(AddTextBox5.Text))));

            PrintWorksheet(ID - int.Parse(RowsCountTextBox.Text) / 2 >= 0 ?
                ID - int.Parse(RowsCountTextBox.Text) / 2 : 0, int.Parse(RowsCountTextBox.Text));
        }

        private void RequestButton4_Click(object sender, RoutedEventArgs e)
        {
            DebugPrint(excelDataBase.HighestAccruedAccountHolder());
        }

        private void RequestButton3_Click(object sender, RoutedEventArgs e)
        {
            DebugPrint(excelDataBase.TheMostLostAccountHolder());
        }

        private void RequestButton2_Click(object sender, RoutedEventArgs e)
        {
            DebugPrint(excelDataBase.IncomeCurrencies());
        }

        private void RequestButton1_Click(object sender, RoutedEventArgs e)
        {
            DebugPrint(excelDataBase.CountAccrualsForAllCur());
        }

        private void NewFileButton_Click(object sender, RoutedEventArgs e)
        {
            AddToProtocol("Clicked NewFileButton");
            Open1();
        }

        private void OldFileButton_Click(object sender, RoutedEventArgs e)
        {


            if (File.Exists(Path.GetFullPath(@"..\..\data\Protocol.txt")))
            {
                using (StreamReader reader = new StreamReader(Path.GetFullPath(@"..\..\data\Protocol.txt")))
                {
                    protocol = reader.ReadToEnd();                   
                }
                AddToProtocol("Clicked OldFileButton");
                AddToProtocol("Try to read old file...");
                AddToProtocol("Success read old file");
            }

            else AddToProtocol("Old file is not found");

            Open1();


        }

        private void Open1()
        {
            AddToProtocol("Try to read Excel file");

            ChangeVisible(false, NewFileButton, OldFileButton);
            ChangeVisible(false, QuestionLabel1);

            try
            {              
                excelDataBase = new ExcelDataBase(Path.GetFullPath(@"..\..\data\LR6-var13.xls"), Path.GetFullPath(@"..\..\data\LR6-var13.xlsx"));
                MessageLabel.Content = "Successful Data Read";

                AddToProtocol("Successful Data Read");
            }
            catch (Exception ex)
            {
                AddToProtocol("Unsuccessful Data Read");
                MessageLabel.Content = ex;
            }

            ChangeVisible(true, ShowWorksheet, delItemButton, CorrectElemButton, AddElemButton, RequestButton1, RequestButton2, RequestButton3, RequestButton4, Reset);
            ChangeVisible(true, startIndexTextBox, RowsCountTextBox, ID_forDel, CorrectIDTextBox, rawItemTextBox, AddTextBox1, AddTextBox2, AddTextBox3, AddTextBox4, AddTextBox5);
            ChangeVisible(true, ColoumsNamesComboBox, WorksheetsNames);
            ChangeVisible(true, OutputLabel, LeftLabel1, LeftLabel2, LeftLabel3, LeftLabel4, LeftLabel5, LeftLabel6, LeftLabel7, LeftLabel8, LeftLabel9, LeftLabel10, LeftLabel11, LeftLabel12, referenceLabel);
        }

        public void AddToProtocol(string text)
        {
            protocol += DateTime.Now.ToString() + " " + text + "\n";
            MessageLabel.Content = DateTime.Now.ToString() + " " + text;
        }

        private void Reset_Click(object sender, RoutedEventArgs e)
        {
            AddToProtocol("Reset was clicked");

            startIndexTextBox.Text = 0.ToString();
            RowsCountTextBox.Text = 10.ToString();

            AddToProtocol("Seccessful reset");
        }
    }
}
