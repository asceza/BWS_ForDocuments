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
            TableData td = new TableData();
            var tableData = new List<TableDataInstance>();

            RangeRow rangeRow = GetRangeRow();


            try
            {
                tableData = td.GetTableData(tbTemplateTablePath.Text, tbTableName.Text);
            }
            catch
            {
                string tName = tbTableName.Text;
                MessageBox.Show(@"В исходном файле не найдена таблица с именем " + tName,
                                 "Ошибка",
                                 MessageBoxButton.OK,
                                 MessageBoxImage.Error);
            }

            CopyDocument cd = new CopyDocument();
            FileNameList fnl = new FileNameList();
            Writer wr = new Writer();

            var fileNames = fnl.GetFileNameList(tableData);

            cd.CopyDoc(fileNames, folderPath, tbTemplateDocPath.Text);

            var destinationFiles = Directory.GetFiles(folderPath).ToList();

            wr.WriteValue(tableData, destinationFiles, tbSymbol.Text, folderPath);
        }

        private RangeRow GetRangeRow()
        {
            try
            {
                if (tbFirstRow.Text == "" && tbLastRow.Text == "")
                {
                    return new RangeRow();
                }

                if (uint.Parse(tbFirstRow.Text) > 0 && tbLastRow.Text == "")
                {
                    return new RangeRow(uint.Parse(tbFirstRow.Text));
                }

                if (uint.Parse(tbFirstRow.Text) > 0 && uint.Parse(tbLastRow.Text) > uint.Parse(tbFirstRow.Text))
                {
                    return new RangeRow(uint.Parse(tbFirstRow.Text), uint.Parse(tbLastRow.Text) - uint.Parse(tbFirstRow.Text));
                }
                return new RangeRow();
            }
            catch (Exception)
            {
                MessageBox.Show(@"Диапазон должен содержать числа больше нуля ",
                                 "Ошибка",
                                 MessageBoxButton.OK,
                                 MessageBoxImage.Error);
                return new RangeRow();
            }
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
    }
}
