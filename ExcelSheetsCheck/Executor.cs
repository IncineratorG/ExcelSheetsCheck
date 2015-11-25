using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.IO;

namespace ExcelSheetsCheck
{
    class Executor
    {
        public static string[] Execute(string pathToReference, string folderWithDocumentsToCheck)
        {
            List<String> errorList = new List<string>();            // Содержит ошибки в работе программы

            Excel.Application ReferenceExcel = null;
            Excel.Workbook ReferenceBook = null;
            Excel.Worksheet ReferenceSheet;

            string[] referenceData = null;

            // ----------- Извлекаем данные из справочника --------------------------------------------------------------
            try
            {
                
                ReferenceExcel = new Excel.Application(); //открыть эксель
                ReferenceBook = ReferenceExcel.Workbooks.Open(pathToReference, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
                ReferenceSheet = (Excel.Worksheet)ReferenceBook.Sheets[1]; //получить 1 лист

                // Заполняем массив данными из справочника
                referenceData = CreateReferenceArray(ReferenceSheet);

                //ReferenceBook.Close(false, Type.Missing, Type.Missing); //закрыть, сохраняя
                ReferenceExcel.Quit(); // выйти из экселя
            }
            catch (Exception e)
            {
                //if (ReferenceBook != null)
                   // ReferenceBook.Close();
                //if (ReferenceExcel != null)
                    ReferenceExcel.Quit();

                string[] exceptionMessage = new string[3];
                exceptionMessage[0] = "FATAL ERROR occured while working with: " + pathToReference + "!";
                errorList.Add(exceptionMessage[0]);
                exceptionMessage[1] = "\n";
                errorList.Add(exceptionMessage[1]);
                exceptionMessage[2] = e.ToString();
                errorList.Add(exceptionMessage[2] + "\n");

                return exceptionMessage;
            }
            // ---------------------------------------------------------------------------------------------------------------


            // Получаем массив имён платёжек, подлежащих проверке
            bool isEmptyFolder = false;
            string[] documentsToCheck = GetDocumentsToCheck(folderWithDocumentsToCheck, out isEmptyFolder);
            errorList.Add("Folder with checking documents is empty: " + isEmptyFolder.ToString() + "\n");

            // Если платёжек в папке нет - завершаем работу.
            // Иначе - проверяем каждую платёжку в папке на совпадение со справочником
            if (isEmptyFolder)
            {
                string[] exceptionMessage = new string[2];
                exceptionMessage[0] = "Folder have no documents to check.";
                errorList.Add(exceptionMessage[0] + "\n");

                return exceptionMessage;
            }
            else
            {
                List<string> resultList = new List<string>();

                for (int i = 0; i < documentsToCheck.Length; ++i)
                {
                    Excel.Application ToCheckExcel = null;
                    Excel.Workbook ToCheckBook = null;
                    Excel.Worksheet ToCheckSheet;

                    // Извлекаем данные из документа, подлежащего проверке и выполняем проверку
                    try
                    {
                        // Формируем имя обрабатываемого в данный момент документа
                        int documentNameStartingIndex = documentsToCheck[i].LastIndexOf("\\") + 1;
                        string documentName = documentsToCheck[i].Substring(documentNameStartingIndex);

                        ToCheckExcel = new Excel.Application(); //открыть эксель
                        ToCheckBook = ToCheckExcel.Workbooks.Open(documentsToCheck[i], Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
                        ToCheckSheet = (Excel.Worksheet)ToCheckBook.Sheets[1]; //получить 1 лист

                        // Извлекаем данные из рассматриваемого документа и
                        // сравниваем их со справочником
                        string[] checkingDocumentData = GetCheckingDocumentData(ToCheckSheet);
                        bool inReference = CompareDocumentWithReference(checkingDocumentData, referenceData);

                        // Формируем результат
                        resultList.Add(documentName);
                        string compareResult;
                        if (inReference)
                            compareResult = "Ok";
                        else
                            compareResult = "Error";
                        resultList.Add(compareResult);

                        //ToCheckBook.Close(false, Type.Missing, Type.Missing); //закрыть, сохраняя
                        ToCheckExcel.Quit(); // выйти из экселя                    
                    }
                    catch (Exception e)
                    {
                        //if (ToCheckBook != null)
                           // ToCheckBook.Close();
                        //if (ToCheckExcel != null)
                            ToCheckExcel.Quit();

                        string[] exceptionMessage = new string[3];
                        exceptionMessage[0] = "FATAL ERROR occured while working with: " + documentsToCheck[i] + "!";
                        exceptionMessage[1] = "\n";
                        exceptionMessage[2] = e.ToString();

                        return exceptionMessage;
                    }
                }

                return resultList.ToArray();
            }

        }

        private static bool CompareDocumentWithReference(string[] checkingDocumentData, string[] referenceData)
        {
            bool inReference = false;

            for (int i = 0; i < referenceData.Length; i = i + 2)
            {
                if (i + 1 >= referenceData.Length)
                    break;
                else
                {
                    if (checkingDocumentData[0].Equals(referenceData[i]) && checkingDocumentData[1].Equals(referenceData[i + 1]))
                    {
                        inReference = true;
                        break;
                    }
                }
            }

            return inReference;
        }

        private static string[] GetCheckingDocumentData(Excel.Worksheet worksheet)
        {
            int rowIndexKPP = 7;//29;
            int columnIndexKPP = 6;//1;

            int rowIndexOKTMO = 27;
            int columnIndexOKTMO = 5;

            string[] checkingDocumentData = new string[2];
            StringBuilder sb = new StringBuilder();

            // Читаем ячейки, содержащие КПП и ОКТМО
            string strWithKPP = worksheet.Cells[rowIndexKPP, columnIndexKPP].Text.ToString();       
            string strWithOKTMO = worksheet.Cells[rowIndexOKTMO, columnIndexOKTMO].Text.ToString();

            // Извлекаем значение КПП из прочитанной ячейки (длина КПП - 9 символов)
            int indexOfKPP = strWithKPP.IndexOf("КПП") + 3;
            int digit = 0;
            while (sb.Length < 9 && indexOfKPP < strWithKPP.Length)
            {
                if (int.TryParse(strWithKPP[indexOfKPP].ToString(), out digit))
                {
                    sb.Append(digit.ToString());
                }
                ++indexOfKPP;
            }

            checkingDocumentData[0] = sb.ToString();
            //System.IO.File.WriteAllText(@"C:\\out.txt", sb.ToString());
            checkingDocumentData[1] = strWithOKTMO.Trim();

            return checkingDocumentData;
        }

        private static string[] GetDocumentsToCheck(string folderWithDocumentsToCheck, out bool isEmptyFolder)
        {
            string[] documentsToCheck = null;
            List<String> listOfDocuments = new List<string>();
            string[] filesInDirectory = null;

            // Если папка существует, читаем имена содержащихся в ней файлов
            if (Directory.Exists(folderWithDocumentsToCheck))
            {
                filesInDirectory = Directory.GetFiles(folderWithDocumentsToCheck);
            }

            // Если в выбранной папке имеются файлы, отбираем из них
            // только те, что имеют расширения ".xls" или ".xlsx"
            if (filesInDirectory != null && filesInDirectory.Length > 0)
            {
                for (int i = 0; i < filesInDirectory.Length; ++i)
                {
                    if (filesInDirectory[i].EndsWith(".xls") || filesInDirectory[i].EndsWith(".xlsx"))
                    {
                        listOfDocuments.Add(filesInDirectory[i]);
                    }
                }
                documentsToCheck = listOfDocuments.ToArray();
            }

            // Проверяем выходной массив на пустоту
            if (documentsToCheck != null && documentsToCheck.Length > 0)
                isEmptyFolder = false;
            else
                isEmptyFolder = true;    

            return documentsToCheck;
        }

        private static string[] CreateReferenceArray(Excel.Worksheet worksheet)
        {
            List<String> dataList = new List<string>();

            int rowIndex = 4;
            int columnIndex = 2;
            int maxRowIndex = rowIndex + 1;

            // Определяем индекс последней заполненной строки
            while (true)
            {
                if (worksheet.Cells[maxRowIndex, 1].Text.Equals(""))
                    break;

                ++maxRowIndex;
            }

            // Извлекаем КПП и ОКТМО
            for (int i = rowIndex; i < maxRowIndex; ++i)
            {
                string dataKPP = worksheet.Cells[i, columnIndex].Text.ToString();
                string dataOKTMO = worksheet.Cells[i, columnIndex + 1].Text.ToString();

                if (dataKPP.Equals("") || dataOKTMO.Equals(""))
                    continue;
                else
                {
                    dataList.Add(dataKPP);                    
                    dataList.Add(dataOKTMO);
                }
            }

            return dataList.ToArray<string>();
        }



/*
        private static bool CompareDocumentWithReference(string[] checkingDocumentData, string[] referenceArray)
        {
            bool inReference = false;

            for (int i = 0; i < referenceArray.Length; i = i + 3)
            {
                if (referenceArray[i].Equals(checkingDocumentData[0]) && 
                    referenceArray[i + 1].Equals(checkingDocumentData[1]) && 
                    referenceArray[i + 2].Equals(checkingDocumentData[2]))
                {
                    inReference = true;
                    break;
                }
            }

            return inReference;
        }

        private static string[] CheckingDocumentData(Excel.Worksheet WorkSheet)
        {
            string[] checkingDocumentData = new string[3];

            int KPP_Row_Index = 29;     // Длинная строка, нужно обработать
            int KPP_Column_Index = 1;

            int OKTMO_Row_Index = 27;       // Не нужно обрабатывать
            int OKTMO_Column_Index = 5;

            int INN_Row_Index = 20;         // ИНН 2005465..
            int INN_Column_Index = 1;

            string KPP = WorkSheet.Cells[KPP_Row_Index, KPP_Column_Index].Text;
            int KPP_Start_Positon = KPP.LastIndexOf("КПП") + 4;
            string KPP_Value = KPP.Substring(KPP_Start_Positon, 9);

            checkingDocumentData[2] = WorkSheet.Cells[INN_Row_Index, INN_Column_Index].Text.Substring(4);
            checkingDocumentData[1] = WorkSheet.Cells[OKTMO_Row_Index, OKTMO_Column_Index].Text;
            checkingDocumentData[0] = KPP_Value;

            return checkingDocumentData;
        }

        private static string[] CreateReferenceArray(Excel.Worksheet WorkSheet)
        {
            List<string> list = new List<string>();
            int rowIndex = 4;
            int columnIndex = 2;
            int maxRowIndex = rowIndex + 1;

            while (true)
            {
                if (WorkSheet.Cells[maxRowIndex, 1].Text.Equals(""))
                    break;

                ++maxRowIndex;
            }

            while (rowIndex < maxRowIndex)
            {
                string value = WorkSheet.Cells[rowIndex, columnIndex].Text.ToString();
                if (value.Equals(""))
                    value = "*";

                list.Add(value);
                ++columnIndex;
                if (columnIndex == 5)
                {
                    ++rowIndex;
                    columnIndex = 2;
                }
            }

            return list.ToArray<string>();
        }
 */ 
    }
}
