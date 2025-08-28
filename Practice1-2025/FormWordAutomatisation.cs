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
            LoadSettings();
            setDefaultParameters();
            UpdateControls();
            //UpdatePreview();
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
