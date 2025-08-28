namespace Practice1_2025
{
    partial class FormWordAutomatisation
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblWorkName = new System.Windows.Forms.Label();
            this.fldWorkName = new System.Windows.Forms.TextBox();
            this.lblTopic = new System.Windows.Forms.Label();
            this.fldTopic = new System.Windows.Forms.TextBox();
            this.lblSubject = new System.Windows.Forms.Label();
            this.fldSubject = new System.Windows.Forms.TextBox();
            this.lblDocumentType = new System.Windows.Forms.Label();
            this.lblWorkType = new System.Windows.Forms.Label();
            this.lblWorkNumber = new System.Windows.Forms.Label();
            this.lblChackerName = new System.Windows.Forms.Label();
            this.lblCheckerRegalies = new System.Windows.Forms.Label();
            this.labellblUniversityName = new System.Windows.Forms.Label();
            this.lblDepartament = new System.Windows.Forms.Label();
            this.lblCity = new System.Windows.Forms.Label();
            this.lblYear = new System.Windows.Forms.Label();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCreate = new System.Windows.Forms.Button();
            this.fldDocumentType = new System.Windows.Forms.ComboBox();
            this.fldWorkType = new System.Windows.Forms.ComboBox();
            this.lblUniversityRegalies = new System.Windows.Forms.Label();
            this.fldCheckerName = new System.Windows.Forms.TextBox();
            this.fldCheckerRegalies = new System.Windows.Forms.TextBox();
            this.fldUniversityName = new System.Windows.Forms.TextBox();
            this.fldUniversityRegalies = new System.Windows.Forms.TextBox();
            this.fldDepartament = new System.Windows.Forms.TextBox();
            this.btnPreview = new System.Windows.Forms.Button();
            this.fldCity = new System.Windows.Forms.TextBox();
            this.fldYear = new System.Windows.Forms.NumericUpDown();
            this.fldWorkNumber = new System.Windows.Forms.NumericUpDown();
            this.fldInstitution = new System.Windows.Forms.TextBox();
            this.lblInstitution = new System.Windows.Forms.Label();
            this.fldGroup = new System.Windows.Forms.TextBox();
            this.lblGroup = new System.Windows.Forms.Label();
            this.fldVariant = new System.Windows.Forms.TextBox();
            this.lblVariant = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.fldYear)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.fldWorkNumber)).BeginInit();
            this.SuspendLayout();
            // 
            // lblWorkName
            // 
            this.lblWorkName.AutoSize = true;
            this.lblWorkName.Location = new System.Drawing.Point(37, 18);
            this.lblWorkName.Name = "lblWorkName";
            this.lblWorkName.Size = new System.Drawing.Size(127, 16);
            this.lblWorkName.TabIndex = 0;
            this.lblWorkName.Text = "Название работы:";
            this.lblWorkName.Click += new System.EventHandler(this.label1_Click);
            // 
            // fldWorkName
            // 
            this.fldWorkName.Location = new System.Drawing.Point(170, 18);
            this.fldWorkName.Name = "fldWorkName";
            this.fldWorkName.Size = new System.Drawing.Size(602, 22);
            this.fldWorkName.TabIndex = 1;
            // 
            // lblTopic
            // 
            this.lblTopic.AutoSize = true;
            this.lblTopic.Location = new System.Drawing.Point(120, 49);
            this.lblTopic.Name = "lblTopic";
            this.lblTopic.Size = new System.Drawing.Size(44, 16);
            this.lblTopic.TabIndex = 2;
            this.lblTopic.Text = "Тема:";
            // 
            // fldTopic
            // 
            this.fldTopic.Location = new System.Drawing.Point(170, 46);
            this.fldTopic.Name = "fldTopic";
            this.fldTopic.Size = new System.Drawing.Size(602, 22);
            this.fldTopic.TabIndex = 3;
            // 
            // lblSubject
            // 
            this.lblSubject.AutoSize = true;
            this.lblSubject.Location = new System.Drawing.Point(74, 84);
            this.lblSubject.Name = "lblSubject";
            this.lblSubject.Size = new System.Drawing.Size(90, 16);
            this.lblSubject.TabIndex = 4;
            this.lblSubject.Text = "Дисциплина:";
            // 
            // fldSubject
            // 
            this.fldSubject.Location = new System.Drawing.Point(170, 81);
            this.fldSubject.Name = "fldSubject";
            this.fldSubject.Size = new System.Drawing.Size(601, 22);
            this.fldSubject.TabIndex = 5;
            // 
            // lblDocumentType
            // 
            this.lblDocumentType.AutoSize = true;
            this.lblDocumentType.Location = new System.Drawing.Point(55, 127);
            this.lblDocumentType.Name = "lblDocumentType";
            this.lblDocumentType.Size = new System.Drawing.Size(109, 16);
            this.lblDocumentType.TabIndex = 6;
            this.lblDocumentType.Text = "Вид документа:";
            // 
            // lblWorkType
            // 
            this.lblWorkType.AutoSize = true;
            this.lblWorkType.Location = new System.Drawing.Point(78, 169);
            this.lblWorkType.Name = "lblWorkType";
            this.lblWorkType.Size = new System.Drawing.Size(86, 16);
            this.lblWorkType.TabIndex = 7;
            this.lblWorkType.Text = "Вид работы:";
            // 
            // lblWorkNumber
            // 
            this.lblWorkNumber.AutoSize = true;
            this.lblWorkNumber.Location = new System.Drawing.Point(61, 213);
            this.lblWorkNumber.Name = "lblWorkNumber";
            this.lblWorkNumber.Size = new System.Drawing.Size(104, 16);
            this.lblWorkNumber.TabIndex = 8;
            this.lblWorkNumber.Text = "Номер работы:";
            // 
            // lblChackerName
            // 
            this.lblChackerName.AutoSize = true;
            this.lblChackerName.Location = new System.Drawing.Point(29, 292);
            this.lblChackerName.Name = "lblChackerName";
            this.lblChackerName.Size = new System.Drawing.Size(135, 16);
            this.lblChackerName.TabIndex = 9;
            this.lblChackerName.Text = "Имя проверяющего:";
            // 
            // lblCheckerRegalies
            // 
            this.lblCheckerRegalies.AutoSize = true;
            this.lblCheckerRegalies.Location = new System.Drawing.Point(-1, 323);
            this.lblCheckerRegalies.Name = "lblCheckerRegalies";
            this.lblCheckerRegalies.Size = new System.Drawing.Size(164, 16);
            this.lblCheckerRegalies.TabIndex = 10;
            this.lblCheckerRegalies.Text = "Регалии проверяющего:";
            this.lblCheckerRegalies.Click += new System.EventHandler(this.lblCheckerRegalies_Click);
            // 
            // labellblUniversityName
            // 
            this.labellblUniversityName.AutoSize = true;
            this.labellblUniversityName.Location = new System.Drawing.Point(52, 448);
            this.labellblUniversityName.Name = "labellblUniversityName";
            this.labellblUniversityName.Size = new System.Drawing.Size(111, 16);
            this.labellblUniversityName.TabIndex = 11;
            this.labellblUniversityName.Text = "Название вуза:";
            // 
            // lblDepartament
            // 
            this.lblDepartament.AutoSize = true;
            this.lblDepartament.Location = new System.Drawing.Point(-5, 524);
            this.lblDepartament.Name = "lblDepartament";
            this.lblDepartament.Size = new System.Drawing.Size(171, 16);
            this.lblDepartament.TabIndex = 12;
            this.lblDepartament.Text = "Наименование кафедры:";
            this.lblDepartament.Click += new System.EventHandler(this.label9_Click);
            // 
            // lblCity
            // 
            this.lblCity.AutoSize = true;
            this.lblCity.Location = new System.Drawing.Point(115, 562);
            this.lblCity.Name = "lblCity";
            this.lblCity.Size = new System.Drawing.Size(49, 16);
            this.lblCity.TabIndex = 13;
            this.lblCity.Text = "Город:";
            // 
            // lblYear
            // 
            this.lblYear.AutoSize = true;
            this.lblYear.Location = new System.Drawing.Point(133, 593);
            this.lblYear.Name = "lblYear";
            this.lblYear.Size = new System.Drawing.Size(33, 16);
            this.lblYear.TabIndex = 14;
            this.lblYear.Text = "Год:";
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(170, 645);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(91, 24);
            this.btnSave.TabIndex = 15;
            this.btnSave.Text = "Сохранить";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnCreate
            // 
            this.btnCreate.Location = new System.Drawing.Point(403, 646);
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.Size = new System.Drawing.Size(92, 23);
            this.btnCreate.TabIndex = 16;
            this.btnCreate.Text = "Создать";
            this.btnCreate.UseVisualStyleBackColor = true;
            this.btnCreate.Click += new System.EventHandler(this.btnCreate_Click);
            // 
            // fldDocumentType
            // 
            this.fldDocumentType.FormattingEnabled = true;
            this.fldDocumentType.Location = new System.Drawing.Point(170, 124);
            this.fldDocumentType.Name = "fldDocumentType";
            this.fldDocumentType.Size = new System.Drawing.Size(601, 24);
            this.fldDocumentType.TabIndex = 17;
            // 
            // fldWorkType
            // 
            this.fldWorkType.FormattingEnabled = true;
            this.fldWorkType.Location = new System.Drawing.Point(170, 166);
            this.fldWorkType.Name = "fldWorkType";
            this.fldWorkType.Size = new System.Drawing.Size(601, 24);
            this.fldWorkType.TabIndex = 18;
            // 
            // lblUniversityRegalies
            // 
            this.lblUniversityRegalies.AutoSize = true;
            this.lblUniversityRegalies.Location = new System.Drawing.Point(63, 486);
            this.lblUniversityRegalies.Name = "lblUniversityRegalies";
            this.lblUniversityRegalies.Size = new System.Drawing.Size(100, 16);
            this.lblUniversityRegalies.TabIndex = 20;
            this.lblUniversityRegalies.Text = "Регалии вуза:";
            this.lblUniversityRegalies.Click += new System.EventHandler(this.label1_Click_1);
            // 
            // fldCheckerName
            // 
            this.fldCheckerName.Location = new System.Drawing.Point(170, 286);
            this.fldCheckerName.Name = "fldCheckerName";
            this.fldCheckerName.Size = new System.Drawing.Size(600, 22);
            this.fldCheckerName.TabIndex = 21;
            // 
            // fldCheckerRegalies
            // 
            this.fldCheckerRegalies.Location = new System.Drawing.Point(169, 323);
            this.fldCheckerRegalies.Name = "fldCheckerRegalies";
            this.fldCheckerRegalies.Size = new System.Drawing.Size(600, 22);
            this.fldCheckerRegalies.TabIndex = 22;
            // 
            // fldUniversityName
            // 
            this.fldUniversityName.Location = new System.Drawing.Point(169, 442);
            this.fldUniversityName.Name = "fldUniversityName";
            this.fldUniversityName.Size = new System.Drawing.Size(600, 22);
            this.fldUniversityName.TabIndex = 23;
            // 
            // fldUniversityRegalies
            // 
            this.fldUniversityRegalies.Location = new System.Drawing.Point(169, 480);
            this.fldUniversityRegalies.Name = "fldUniversityRegalies";
            this.fldUniversityRegalies.Size = new System.Drawing.Size(600, 22);
            this.fldUniversityRegalies.TabIndex = 24;
            // 
            // fldDepartament
            // 
            this.fldDepartament.Location = new System.Drawing.Point(169, 518);
            this.fldDepartament.Name = "fldDepartament";
            this.fldDepartament.Size = new System.Drawing.Size(600, 22);
            this.fldDepartament.TabIndex = 25;
            // 
            // btnPreview
            // 
            this.btnPreview.Location = new System.Drawing.Point(636, 645);
            this.btnPreview.Name = "btnPreview";
            this.btnPreview.Size = new System.Drawing.Size(134, 23);
            this.btnPreview.TabIndex = 26;
            this.btnPreview.Text = "Предпросмотр";
            this.btnPreview.UseVisualStyleBackColor = true;
            this.btnPreview.Click += new System.EventHandler(this.btnPreview_Click);
            // 
            // fldCity
            // 
            this.fldCity.Location = new System.Drawing.Point(170, 556);
            this.fldCity.Name = "fldCity";
            this.fldCity.Size = new System.Drawing.Size(600, 22);
            this.fldCity.TabIndex = 27;
            // 
            // fldYear
            // 
            this.fldYear.Location = new System.Drawing.Point(172, 593);
            this.fldYear.Maximum = new decimal(new int[] {
            2100,
            0,
            0,
            0});
            this.fldYear.Minimum = new decimal(new int[] {
            1900,
            0,
            0,
            0});
            this.fldYear.Name = "fldYear";
            this.fldYear.Size = new System.Drawing.Size(597, 22);
            this.fldYear.TabIndex = 28;
            this.fldYear.Value = new decimal(new int[] {
            2025,
            0,
            0,
            0});
            this.fldYear.ValueChanged += new System.EventHandler(this.fldYear_ValueChanged);
            // 
            // fldWorkNumber
            // 
            this.fldWorkNumber.Location = new System.Drawing.Point(170, 207);
            this.fldWorkNumber.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.fldWorkNumber.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.fldWorkNumber.Name = "fldWorkNumber";
            this.fldWorkNumber.Size = new System.Drawing.Size(601, 22);
            this.fldWorkNumber.TabIndex = 29;
            this.fldWorkNumber.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // fldInstitution
            // 
            this.fldInstitution.Location = new System.Drawing.Point(169, 400);
            this.fldInstitution.Name = "fldInstitution";
            this.fldInstitution.Size = new System.Drawing.Size(600, 22);
            this.fldInstitution.TabIndex = 30;
            // 
            // lblInstitution
            // 
            this.lblInstitution.AutoSize = true;
            this.lblInstitution.Location = new System.Drawing.Point(92, 406);
            this.lblInstitution.Name = "lblInstitution";
            this.lblInstitution.Size = new System.Drawing.Size(72, 16);
            this.lblInstitution.TabIndex = 31;
            this.lblInstitution.Text = "Институт:";
            // 
            // fldGroup
            // 
            this.fldGroup.Location = new System.Drawing.Point(169, 361);
            this.fldGroup.Name = "fldGroup";
            this.fldGroup.Size = new System.Drawing.Size(600, 22);
            this.fldGroup.TabIndex = 32;
            // 
            // lblGroup
            // 
            this.lblGroup.AutoSize = true;
            this.lblGroup.Location = new System.Drawing.Point(50, 367);
            this.lblGroup.Name = "lblGroup";
            this.lblGroup.Size = new System.Drawing.Size(115, 16);
            this.lblGroup.TabIndex = 33;
            this.lblGroup.Text = "Учебная группа:";
            // 
            // fldVariant
            // 
            this.fldVariant.Location = new System.Drawing.Point(169, 247);
            this.fldVariant.Name = "fldVariant";
            this.fldVariant.Size = new System.Drawing.Size(600, 22);
            this.fldVariant.TabIndex = 34;
            // 
            // lblVariant
            // 
            this.lblVariant.AutoSize = true;
            this.lblVariant.Location = new System.Drawing.Point(31, 253);
            this.lblVariant.Name = "lblVariant";
            this.lblVariant.Size = new System.Drawing.Size(119, 16);
            this.lblVariant.TabIndex = 35;
            this.lblVariant.Text = "Номер варианта:";
            // 
            // FormWordAutomatisation
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 680);
            this.Controls.Add(this.lblVariant);
            this.Controls.Add(this.fldVariant);
            this.Controls.Add(this.lblGroup);
            this.Controls.Add(this.fldGroup);
            this.Controls.Add(this.lblInstitution);
            this.Controls.Add(this.fldInstitution);
            this.Controls.Add(this.fldWorkNumber);
            this.Controls.Add(this.fldYear);
            this.Controls.Add(this.fldCity);
            this.Controls.Add(this.btnPreview);
            this.Controls.Add(this.fldDepartament);
            this.Controls.Add(this.fldUniversityRegalies);
            this.Controls.Add(this.fldUniversityName);
            this.Controls.Add(this.fldCheckerRegalies);
            this.Controls.Add(this.fldCheckerName);
            this.Controls.Add(this.lblUniversityRegalies);
            this.Controls.Add(this.fldWorkType);
            this.Controls.Add(this.fldDocumentType);
            this.Controls.Add(this.btnCreate);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.lblYear);
            this.Controls.Add(this.lblCity);
            this.Controls.Add(this.lblDepartament);
            this.Controls.Add(this.labellblUniversityName);
            this.Controls.Add(this.lblCheckerRegalies);
            this.Controls.Add(this.lblChackerName);
            this.Controls.Add(this.lblWorkNumber);
            this.Controls.Add(this.lblWorkType);
            this.Controls.Add(this.lblDocumentType);
            this.Controls.Add(this.fldSubject);
            this.Controls.Add(this.lblSubject);
            this.Controls.Add(this.fldTopic);
            this.Controls.Add(this.lblTopic);
            this.Controls.Add(this.fldWorkName);
            this.Controls.Add(this.lblWorkName);
            this.Name = "FormWordAutomatisation";
            this.Text = "Кручинин И.Н.; Номер варианта: 6";
            this.Load += new System.EventHandler(this.FormWordAutomatisation_Load);
            ((System.ComponentModel.ISupportInitialize)(this.fldYear)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.fldWorkNumber)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblWorkName;
        private System.Windows.Forms.TextBox fldWorkName;
        private System.Windows.Forms.Label lblTopic;
        private System.Windows.Forms.TextBox fldTopic;
        private System.Windows.Forms.Label lblSubject;
        private System.Windows.Forms.TextBox fldSubject;
        private System.Windows.Forms.Label lblDocumentType;
        private System.Windows.Forms.Label lblWorkType;
        private System.Windows.Forms.Label lblWorkNumber;
        private System.Windows.Forms.Label lblChackerName;
        private System.Windows.Forms.Label lblCheckerRegalies;
        private System.Windows.Forms.Label labellblUniversityName;
        private System.Windows.Forms.Label lblDepartament;
        private System.Windows.Forms.Label lblCity;
        private System.Windows.Forms.Label lblYear;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnCreate;
        private System.Windows.Forms.ComboBox fldDocumentType;
        private System.Windows.Forms.ComboBox fldWorkType;
        private System.Windows.Forms.Label lblUniversityRegalies;
        private System.Windows.Forms.TextBox fldCheckerName;
        private System.Windows.Forms.TextBox fldCheckerRegalies;
        private System.Windows.Forms.TextBox fldUniversityName;
        private System.Windows.Forms.TextBox fldUniversityRegalies;
        private System.Windows.Forms.TextBox fldDepartament;
        private System.Windows.Forms.Button btnPreview;
        private System.Windows.Forms.TextBox fldCity;
        private System.Windows.Forms.NumericUpDown fldYear;
        private System.Windows.Forms.NumericUpDown fldWorkNumber;
        private System.Windows.Forms.TextBox fldInstitution;
        private System.Windows.Forms.Label lblInstitution;
        private System.Windows.Forms.TextBox fldGroup;
        private System.Windows.Forms.Label lblGroup;
        private System.Windows.Forms.TextBox fldVariant;
        private System.Windows.Forms.Label lblVariant;
    }
}

