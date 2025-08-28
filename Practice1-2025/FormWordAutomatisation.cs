using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace Practice1_2025
{
    public partial class FormWordAutomatisation : Form
    {
        private string settingsFile = "settings.txt";
        private string[] parameters = new string[19]; // 19 параметров по списку
        private string authorName;


        public FormWordAutomatisation()
        {
            InitializeComponent();
            authorName = this.Text; // Например, имя автора задано в свойствах формы как "Фамилия И.О."
            setDefaultParameters();
            LoadSettings();
            UpdateControls();
            initializeComboBoxes();
            //UpdatePreview();
        }

        private void initializeComboBoxes()
        {

            // Заполняем ComboBox для вида документа
            var documentTypes = new List<string>
            {
                "отчёт",
                "реферат",
                "эссе",
                "курсовой проект",
                "курсовая работа",
                "доклад",
                "домашнее задание"

            };

            // Заполняем ComboBox для вида работы
            var workTypes = new List<string>
            {
                "Учебная практика",
                "Производственная практика",
                "Лабораторная работа",
                "Практическомое занятие",
                "Индивидуальное задание"
            };

            fldDocumentType.Items.AddRange(documentTypes.ToArray());
            fldWorkType.Items.AddRange(workTypes.ToArray());
        }

        private void LoadSettings()
        {
            if (File.Exists(settingsFile))
            {
                string[] lines = File.ReadAllLines(settingsFile);
                if (lines.Length > 0)
                {
                    for (int i = 0; i < lines.Length && i < parameters.Length; i++)
                    {
                        parameters[i] = lines[i].Trim();
                    }
                }
            }
        }

        private void setDefaultParameters()
        {
            parameters[0] = "Министерство транспорта Российской Федерации";
            parameters[1] = "Федеральное государственное автономное образовательное учреждение высшего образования";
            parameters[2] = "«Российский университет транспорта» (РУТ (МИИТ))";
            parameters[3] = "Институт транспортной техники и систем управления";
            parameters[4] = "Кафедра «Управление и защита информации»";
            parameters[5] = "Отчёт";
            parameters[6] = "Учебной практике";
            parameters[7] = "1";
            parameters[8] = "";
            parameters[9] = "Языки программирования";
            parameters[10] = "Автоматизация Word";
            parameters[11] = "ТКИ-341";
            parameters[12] = authorName;
            parameters[13] = "1";
            parameters[14] = "Сафронов А.И.";
            parameters[15] = "доц. каф. УИЗИ ИТТСУ, к.т.н.";
            parameters[16] = "Москва";
            parameters[17] = "2025";

        }

        private void UpdateControls()
        {
            // Заполняем элементы управления
            fldUniversityRegalies.Text = parameters[1];
            fldUniversityName.Text = parameters[2];
            fldInstitution.Text = parameters[3];
            fldDepartament.Text = parameters[4];

            fldDocumentType.Text = parameters[5];
            fldWorkType.Text = parameters[6];
            fldWorkNumber.Text = parameters[7];
            fldWorkName.Text = parameters[8];
            fldSubject.Text = parameters[9];
            fldTopic.Text = parameters[10];
            fldGroup.Text = parameters[11];
            fldVariant.Text = parameters[13];
            fldCheckerName.Text = parameters[14];
            fldCheckerRegalies.Text = parameters[15];
            fldCity.Text = parameters[16];
            fldYear.Text = parameters[17];

       
        }


        private void updateParametersFromControls()
        {
            parameters[1] = fldUniversityRegalies.Text;
            parameters[2] = fldUniversityName.Text;
            parameters[3] = fldInstitution.Text;
            parameters[4] = fldDepartament.Text;
            parameters[5] = fldDocumentType.Text;
            // Тип работы нужно склонять, пока его не захватываем
            parameters[7] = fldWorkNumber.Text;
            parameters[8] = fldWorkName.Text;
            parameters[9] = fldSubject.Text;
            parameters[10] = fldTopic.Text;
            parameters[11] = fldGroup.Text;
            parameters[12] = authorName;
            parameters[13] = fldVariant.Text;
            parameters[14] = fldCheckerName.Text;
            parameters[15] = fldCheckerRegalies.Text;
            parameters[16] = fldCity.Text;
            parameters[17] = fldYear.Text;

            // Склоняем вид работы (по + сущ. в Д. П.)
            string workType = fldWorkType.Text;
            switch (workType)
            {
                case "Лабораторная работа":
                    workType = "Лабораторной работе";
                    break;
                case "Практическая работа":
                    workType = "Практической работе";
                    break;
                case "Индивидуальное задание":
                    workType = "Индивидуальному заданию";
                    break;
                case "Учебная практика":
                    workType = "Учебной практике";
                    break;
                case "Производственная практика":
                    workType = "Производственной практике";
                    break;
                case "Преддипломная практика":
                    workType = "Преддипломной практике";
                    break;
                default:
                    workType = null;
                    break;
            }
            parameters[6] = workType;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            updateParametersFromControls();
            File.WriteAllLines(settingsFile, parameters.Take(18).Where(p => p != null).ToArray());
            MessageBox.Show("Настройки сохранены!", "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void createWordDocument(string filename)
        {
            updateParametersFromControls();

            Word.Application wordApp = null;
            Word.Document wordDoc = null;

            try
            {
                wordApp = new Word.Application();
                wordDoc = wordApp.Documents.Add();

                Word.Paragraph para;

                object oEndOfDoc = "\\endofdoc"; // Конец документа

                // Установка шрифта по умолчанию
                wordDoc.Content.Font.Name = "Times New Roman";
                wordDoc.Content.Font.Size = 14;

                // 1. Министерство
                para = wordDoc.Content.Paragraphs.Add();
                para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para.Range.Text = parameters[0];
                para.Format.SpaceAfter = 6;
                para.Range.InsertParagraphAfter();

                // 2. Регалии вуза
                para = wordDoc.Content.Paragraphs.Add();
                para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para.Range.Text = parameters[1];
                para.Format.SpaceAfter = 6;
                para.Range.InsertParagraphAfter();

                // 3. Название вуза
                para = wordDoc.Content.Paragraphs.Add();
                para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para.Range.Text = parameters[2];
                para.Format.SpaceAfter = 12;
                para.Range.InsertParagraphAfter();

                // 4. Институт
                para = wordDoc.Content.Paragraphs.Add();
                para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para.Range.Text = parameters[3];
                para.Format.SpaceAfter = 6;
                para.Range.InsertParagraphAfter();

                // 5. Кафедра
                para = wordDoc.Content.Paragraphs.Add();
                para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para.Range.Text = parameters[4];
                para.Format.SpaceAfter = 24;
                para.Range.InsertParagraphAfter();

                // 6. Вид документа
                para = wordDoc.Content.Paragraphs.Add();
                para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para.Range.Text = parameters[5];
                para.Range.Font.Bold = 1;
                para.Format.SpaceAfter = 6;
                para.Range.InsertParagraphAfter();

                // 7. Вид занятия
                para = wordDoc.Content.Paragraphs.Add();
                para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para.Range.Text = $"по {parameters[6]}";
                para.Range.Font.Bold = 1;
                para.Format.SpaceAfter = 6;
                para.Range.InsertParagraphAfter();

                // 8. Номер работы
                if (!string.IsNullOrEmpty(parameters[7]) && !string.IsNullOrEmpty(parameters[8]))
                {
                    para = wordDoc.Content.Paragraphs.Add();
                    para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    para.Range.Text = $"Выполнена работа №{parameters[7]} на тему: \"{parameters[8]}\"";
                    para.Range.Font.Bold = 1;
                    para.Format.SpaceAfter = 24;
                    para.Range.InsertParagraphAfter();
                }
                else if (!string.IsNullOrEmpty(parameters[8]))
                {
                    para = wordDoc.Content.Paragraphs.Add();
                    para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    para.Range.Text = $"на тему: \"{parameters[8]}\"";
                    para.Range.Font.Bold = 1;
                    para.Format.SpaceAfter = 24;
                    para.Range.InsertParagraphAfter();
                }

                // 9. Дисциплина
                para = wordDoc.Content.Paragraphs.Add();
                para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para.Range.Text = $"по дисциплине \"{parameters[9]}\"";
                para.Format.SpaceAfter = 36;
                para.Range.InsertParagraphAfter();

                // 10. Тема
                para = wordDoc.Content.Paragraphs.Add();
                para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para.Range.Text = $"\"{parameters[10]}\"";
                para.Range.Font.Italic = 1;
                para.Format.SpaceAfter = 36;
                para.Range.InsertParagraphAfter();

                // 11. Таблица для подписей (без границ)
                Word.Table signatureTable = wordDoc.Content.Tables.Add(wordDoc.Bookmarks.get_Item(ref oEndOfDoc).Range, 2, 2);
                signatureTable.Columns[1].SetWidth(200, Word.WdRulerStyle.wdAdjustNone);
                signatureTable.Columns[2].SetWidth(200, Word.WdRulerStyle.wdAdjustNone);
                signatureTable.Borders.Enable = 0; // Нет границ

                // Выполнил
                signatureTable.Cell(1, 1).Range.Text = "Выполнил: ст. гр. " + parameters[11];
                signatureTable.Cell(1, 2).Range.Text = parameters[12];

                // Проверил
                signatureTable.Cell(2, 1).Range.Text = "Проверил: " + parameters[15];
                signatureTable.Cell(2, 2).Range.Text = parameters[14];

                // Добавляем отступ
                wordDoc.Content.InsertParagraphAfter();

                // Город – год
                para = wordDoc.Content.Paragraphs.Add();
                para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                para.Range.Text = $"{parameters[16]} – {parameters[17]} г.";
                para.Format.SpaceAfter = 12;
                para.Range.InsertParagraphAfter();

                // Сохраняем документ
                string docPath = Path.Combine(Application.StartupPath, filename);
                wordDoc.SaveAs2(docPath);
                wordDoc.Close();
                wordApp.Quit();

                MessageBox.Show($"Документ сохранён: {docPath}", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при создании документа: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (wordDoc != null) wordDoc.Close();
                if (wordApp != null) wordApp.Quit();
            }

        }

        private void showPreviewInWord(string tempFilePath)
        {
            Word.Application wordApp = null;
            Word.Document wordDoc = null;
            try {
                wordApp = new Word.Application();
                wordDoc = wordApp.Documents.Add();
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = tempFilePath,
                    UseShellExecute = true
                });

                MessageBox.Show($"Документ открыт в Word.\nФайл: {tempFilePath}",
                    "Предварительный просмотр", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при создании или открытии документа: " + ex.Message,
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                wordDoc?.Close();
                wordApp?.Quit();
    }
}

        private void btnCreate_Click(object sender, EventArgs e)
        {
            createWordDocument("titlePage.docx");
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            createWordDocument("preview.docx");
            showPreviewInWord("preview.docx");
        }

        private void FormWordAutomatisation_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void lblCheckerRegalies_Click(object sender, EventArgs e)
        {

        }

        private void fldYear_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
