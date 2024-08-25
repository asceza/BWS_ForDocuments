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
       
        private void btTamplateSource_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx|All Files|*.*";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Title = "Выберите исходную таблицу";

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                tbTamplateTablePath.Text = filePath;
            }

        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
           
            if (tbTamplateTablePath.Text == "Выберите исходную таблицу" || tbTamplateDocPath.Text == "Выберите шаблон документа")
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
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            if (ms == MessageBoxResult.OK)
            {
                MainFunc(folderPath);
            }

            MessageBoxResult ms2 = MessageBox.Show("Размножение документов завершено\n\nМожете проверять результат\n\nЗакрыть программу?",
                "Процесс завершен",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);
            if (ms == MessageBoxResult.Yes)
            {
                Close();
            }
            

        }

        public void MainFunc(string folderPath)
        {
            TableData td = new TableData();            
            var tableData = new List <TableDataInstance>();

            try
            {
                tableData = td.GetTableData(tbTamplateTablePath.Text, tbTableName.Text);
            }
            catch 
            {
                string tName = tbTableName.Text;
                MessageBoxResult msb = MessageBox.Show(@"В исходном файле не найдена таблица с именем "+tName,
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
            }
            
            CopyDocument cd = new CopyDocument();
            FileNameList fnl = new FileNameList();
            Writer wr = new Writer();

            var fileNames = fnl.GetFileNameList(tableData);

            cd.CopyDoc(fileNames, folderPath, tbTamplateDocPath.Text);

            var destinationFiles = Directory.GetFiles(folderPath).ToList();

            wr.WriteValue(tableData, destinationFiles, tbSymbol.Text, folderPath);
        }

        private void btTamplateDocSource_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Title = "Выберите шаблон документа";

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                tbTamplateDocPath.Text = filePath;
            }

        }

    
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {

        }

        
    }
}
