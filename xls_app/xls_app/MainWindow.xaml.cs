using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.WindowsAPICodePack.Dialogs;
using DocumentFormat.OpenXml.Drawing;
using System.Threading;
using System.Windows.Threading;
using System.Diagnostics;
using xls_app.Properties;


namespace xls_app
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btTemplateSource_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx|All Files|*.*";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Title = "Выберите исходную таблицу";

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                tbTemplateTablePath.Text = filePath;
            }
        }

        private async void btDocMultiply_Click(object sender, RoutedEventArgs e)
        {

            if (tbTemplateTablePath.Text == "Выберите исходную таблицу" || tbTemplateDocPath.Text == "Выберите шаблон документа")
            {
                MessageBoxResult msb = MessageBox.Show(@"Выберите исходную таблицу и/или шаблон документа",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            string folderPath = "";

            var dialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                Title = "Выберите папку куда сохранятся документы"
            };

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                folderPath = dialog.FileName + @"\";
            }
            MessageBoxResult ms = MessageBox.Show("Размножение документов может занять некоторое время\n\nДождитесь сообщения о завершении размножения документов",
                "Сообщение",
                MessageBoxButton.OKCancel,
                MessageBoxImage.Information);
            if (ms == MessageBoxResult.OK)
            {
                MainFunc(folderPath);
                MessageBoxResult ms2 = MessageBox.Show("Размножение документов завершено\n\nМожете проверять результат",
                "Процесс завершен");
            }
        }

        public void MainFunc(string folderPath)
        {
            TableData tableData = new TableData();
            var tableDataList = new List<TableDataInstance>();

            RangeRow rangeRow = GetRangeRow();

            try
            {
                switch (rangeRow.rangeState)
                {
                    case RangeRow.State.Single:
                        tableDataList = tableData.GetTableData(tbTemplateTablePath.Text, tbTableName.Text, rangeRow.FirstRow);
                        break;
                    case RangeRow.State.Several:
                        tableDataList = tableData.GetTableData(tbTemplateTablePath.Text, tbTableName.Text, rangeRow.FirstRow, rangeRow.LastRow);
                        break;
                    case RangeRow.State.All:
                        tableDataList = tableData.GetTableData(tbTemplateTablePath.Text, tbTableName.Text);
                        break;
                    default:
                        break;
                }

            }
            catch
            {
                string tName = tbTableName.Text;
                MessageBox.Show(@"К файлу нет доступа или в исходном файле не найдена таблица с именем " + tName,
                                 "Ошибка",
                                 MessageBoxButton.OK,
                                 MessageBoxImage.Error);
            }

            CopyDocument cd = new CopyDocument();
            FileNameList fnl = new FileNameList();
            Writer wr = new Writer();

            var fileNames = fnl.GetFileNameList(tableDataList);

            cd.CopyDoc(fileNames, folderPath, tbTemplateDocPath.Text);

            var destinationFiles = Directory.GetFiles(folderPath).ToList();

            wr.WriteValue(tableDataList, destinationFiles, tbSymbol.Text, folderPath);
        }

        /// <summary>
        /// Определение диапазона строк по данным из полей ввода<br/>
        /// TODO: учесть все варианты ввода
        /// </summary>
        /// <returns></returns>
        private RangeRow GetRangeRow()
        {

            if (uint.Parse(tbFirstRow.Text) > 0 && tbLastRow.Text == "")
            {
                return new RangeRow(uint.Parse(tbFirstRow.Text));
            }

            if (uint.Parse(tbFirstRow.Text) > 0 && uint.Parse(tbLastRow.Text) > uint.Parse(tbFirstRow.Text))
            {
                return new RangeRow(uint.Parse(tbFirstRow.Text), uint.Parse(tbLastRow.Text));
            }

            return new RangeRow();
        }

        private void btTemplateDocSource_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Title = "Выберите шаблон документа";

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                tbTemplateDocPath.Text = filePath;
            }

        }

        /// <summary>
        /// Обработки нажатия кнопки "Инструкция"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btInstruction_Click(object sender, RoutedEventArgs e)
        {
            string donatUrl = "https://vk.com/video-211694366_456239091";
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = donatUrl,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        /// <summary>
        /// Обработка кнопки "Благодарность"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btGratitude_Click(object sender, RoutedEventArgs e)
        {
            string donatUrl = "https://pay.market-tips.kontur.ru/pay/5221/";
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = donatUrl,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось открыть сайт: " + ex.Message + "\n\n\n Воспользуйтесь QR-кодом");
                HelpDonat helpDonat = new HelpDonat();
                helpDonat.ShowDialog();
            }

        }

        /// <summary>
        /// Обработка изменения содержимого в текстовых полях tbFirstRow и tbLastRow
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RangeRow_TextChanged(object sender, TextChangedEventArgs e)
        {
            uint contentFirstRow = 88;
            uint contentLastRow;
            if (((uint.TryParse(tbFirstRow.Text, out contentFirstRow) && (contentFirstRow > 0)) || (tbFirstRow.Text == ""))
              && ((uint.TryParse(tbLastRow.Text, out contentLastRow) && (contentLastRow > contentFirstRow)) || (tbLastRow.Text == "")))
            {
                tbFirstRow.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
                tbLastRow.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
                btDocMultiply.IsEnabled = true;

            }
            else
            {
                btDocMultiply.IsEnabled = false;
                tbFirstRow.Background = new SolidColorBrush(Color.FromRgb(250, 161, 155));
                tbLastRow.Background = new SolidColorBrush(Color.FromRgb(250, 161, 155));
            }

        }
    }
}
