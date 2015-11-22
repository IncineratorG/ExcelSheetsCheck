using System;
using System.Collections.Generic;
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
using System.Windows.Forms;
using System.IO;

using System.ComponentModel;    // BackgroundWorker

namespace ExcelSheetsCheck
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string referencePath;
        private bool referencePathSet = false;
        private string folderWithDocumentsToCheck;
        private bool folderWithDocumentsToCheckSet = false;
        private BackgroundWorker bw;

        public MainWindow()
        {
            InitializeComponent();

            InitButtons();
            bw = new BackgroundWorker();
            bw.DoWork += new DoWorkEventHandler(StartExecutor);
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(ExecutorCompletedWork);
        }

       

        // Пока не выбраны справочник и документы для проверки, запустить программу нельзя
        private void InitButtons()
        {
            executeBtn.IsEnabled = false;
        }

        // Действие при нажатии кнопки "Отмена"
        private void Cancel(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        // Указываем путь до справочника
        private void SelectReference(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel files (*.xls; *.xlsx)|*.xls;*.xlsx";        // Только файлы .xls и .xlsx
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                FileInfo fi = new FileInfo(ofd.FileName);
                referencePath = fi.FullName;

                // Выводим имя справочника
                if (referencePath != null && !referencePath.Equals(""))
                {
                    referencePathSet = true;
                    textBlockReference.Text = "Справочник: " + referencePath.Substring(referencePath.LastIndexOf("\\") + 1) + "\n";
                }

                if (referencePathSet && folderWithDocumentsToCheckSet)
                    executeBtn.IsEnabled = true;
            } 
        }

        // Выбираем папку с платёжками
        private void SelectFolderWithDocumentsToCheck(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            folderWithDocumentsToCheck = dialog.SelectedPath;

            // Выводим имя папки с платёжками
            if (folderWithDocumentsToCheck != null && !folderWithDocumentsToCheck.Equals(""))
            {
                folderWithDocumentsToCheckSet = true;
                textBlockDocuments.Text = "Папка с платёжками: " + folderWithDocumentsToCheck;
            }

            if (referencePathSet && folderWithDocumentsToCheckSet)
                executeBtn.IsEnabled = true;
        }


        // При нажатии на кнопку выполнить
        private void Execute(object sender, RoutedEventArgs e)
        {
            textBox.Clear();

            if (!bw.IsBusy)
            {
                bw.RunWorkerAsync();
                executeBtn.IsEnabled = false;
            }
        }


        // ------------- Работа BackgroundWorker ----------------------------
        private void StartExecutor(object sender, DoWorkEventArgs e)
        {
            string[] executorResults = ExcelSheetsCheck.Executor.Execute(referencePath, folderWithDocumentsToCheck);

            e.Result = executorResults;
        }

        private void ExecutorCompletedWork(object sender, RunWorkerCompletedEventArgs e)
        {
            string[] executorResults = (string[])e.Result;

            if (executorResults != null && executorResults.Length != 0)
            {
                for (int i = 0; i < executorResults.Length; i = i + 2)
                {
                    if (i + 1 >= executorResults.Length)
                    {
                        textBox.AppendText(executorResults[i]);
                        break;
                    }
                    else
                    {
                        textBox.AppendText(executorResults[i] + "    " + executorResults[i + 1] + "\n");
                    }
                }
            }
            else
            {
                textBox.AppendText("FATAL ERROR! \"referenceData\" is NULL or EMPTY.");
            }

            executeBtn.IsEnabled = true;
        }
        // -----------------------------------------------------------------------
        // Выполняем проверку платёжек
 /*       private void StartExecutor(object sender, RoutedEventArgs e)
        {
            textBox.Clear();

            string[] referenceData = ExcelSheetsCheck.Executor.Execute(referencePath, folderWithDocumentsToCheck);

            if (referenceData != null && referenceData.Length != 0)
            {
                for (int i = 0; i < referenceData.Length; i = i + 2)
                {
                    if (i + 1 >= referenceData.Length)
                    {
                        textBox.AppendText(referenceData[i]);
                        break;
                    }
                    else
                    {
                        textBox.AppendText(referenceData[i] + "    " + referenceData[i + 1] + "\n");
                    }
                }
            }
            else
            {
                textBox.AppendText("FATAL ERROR! \"referenceData\" is NULL or EMPTY.");
            }

        }
*/

    }
}
