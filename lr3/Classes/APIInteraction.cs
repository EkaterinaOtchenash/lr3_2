using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace lr3.Classes
{
    public class APIInteraction
    {
        private string fullName;

        // Метод для получения ФИО через API
        public string GetFullName()
        {
            string URL = "http://localhost:4444/TransferSimulator/fullName";
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(URL);
            request.Method = "GET";

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream());
            string text = reader.ReadToEnd();

            JObject jObject = JObject.Parse(text);
            string value = (string)jObject["value"];
            fullName = value;
            return fullName;
        }

        // Метод для проверки ФИО на запрещённые символы
        public bool ContainsExtraChars(string fullName)
        {
            // Регулярное выражение для проверки на наличие запрещённых символов
            string pattern = @"[^а-яА-Я\s]"; // Разрешены только русские буквы и пробелы
            return Regex.IsMatch(fullName, pattern);
        }

        // Метод для заполнения документа данными
        public string FillDocument()
        {
            if (fullName == null)
            {
                MessageBox.Show("Данные не были получены");
                return "";
            }

            bool isValidFullName = !ContainsExtraChars(fullName);
            string result = isValidFullName ? "ФИО не содержит запрещенные символы" : "ФИО содержит запрещенные символы";

            string[] rowData = { $"Введены данные \n{fullName}", result, isValidFullName ? "Успешно" : "Не успешно" };
            AddToWordTable(rowData);

            return result;
        }

        // Метод для добавления данных в таблицу Word-документа
        public void AddToWordTable(string[] rowData)
        {
            var openFileDlg = new System.Windows.Forms.OpenFileDialog();
            string filePath = "";

            if (openFileDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                filePath = openFileDlg.FileName;
            else
                return;

            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;

            try
            {
                doc = wordApp.Documents.Open(filePath);
                wordApp.Visible = false;

                Word.Table table = doc.Tables[1];
                Word.Row row = table.Rows.Add();

                for (int i = 0; i < rowData.Length; i++)
                {
                    row.Cells[i + 1].Range.Text = rowData[i];
                }

                doc.Save();
                MessageBox.Show("Информация была добавлена в файл!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при работе с документом: " + ex.Message);
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close(Word.WdSaveOptions.wdSaveChanges);
                }
                wordApp.Quit(Word.WdSaveOptions.wdSaveChanges);
            }
        }
    }
}