namespace FormationOfSchedule
{
    partial class ReportDate
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
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btn_OK = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.listBoxPeriods = new DevExpress.XtraEditors.ListBoxControl();
            this.label5 = new System.Windows.Forms.Label();
            this.gridControl1 = new DevExpress.XtraGrid.GridControl();
            this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.PFMNameColumn = new DevExpress.XtraGrid.Columns.GridColumn();
            this.PartnerTypeColumn = new DevExpress.XtraGrid.Columns.GridColumn();
            this.FinPositionEPLColumn = new DevExpress.XtraGrid.Columns.GridColumn();
            this.SummColumn = new DevExpress.XtraGrid.Columns.GridColumn();
            this.DateStartColumn = new DevExpress.XtraGrid.Columns.GridColumn();
            this.DateEndColumn = new DevExpress.XtraGrid.Columns.GridColumn();
            this.label6 = new System.Windows.Forms.Label();
            this.btn_ShowDetail = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.listBoxPeriods)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(48, 39);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(200, 20);
            this.dateTimePicker1.TabIndex = 0;
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.Location = new System.Drawing.Point(48, 76);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(200, 20);
            this.dateTimePicker2.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.ForeColor = System.Drawing.SystemColors.Desktop;
            this.label1.Location = new System.Drawing.Point(32, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(206, 17);
            this.label1.TabIndex = 1;
            this.label1.Text = "Введите период платежей";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.ForeColor = System.Drawing.SystemColors.Desktop;
            this.label2.Location = new System.Drawing.Point(29, 40);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(16, 17);
            this.label2.TabIndex = 1;
            this.label2.Text = "с";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.ForeColor = System.Drawing.SystemColors.Desktop;
            this.label3.Location = new System.Drawing.Point(19, 77);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(26, 17);
            this.label3.TabIndex = 1;
            this.label3.Text = "по";
            // 
            // btn_OK
            // 
            this.btn_OK.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btn_OK.Location = new System.Drawing.Point(62, 114);
            this.btn_OK.Name = "btn_OK";
            this.btn_OK.Size = new System.Drawing.Size(159, 34);
            this.btn_OK.TabIndex = 1;
            this.btn_OK.Text = "Показать отчет";
            this.btn_OK.UseVisualStyleBackColor = true;
            this.btn_OK.Click += new System.EventHandler(this.btn_OK_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label4.Location = new System.Drawing.Point(277, 9);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(259, 16);
            this.label4.TabIndex = 3;
            this.label4.Text = "Сохраненные данные по декадам:";
            // 
            // listBoxPeriods
            // 
            this.listBoxPeriods.Location = new System.Drawing.Point(280, 55);
            this.listBoxPeriods.Name = "listBoxPeriods";
            this.listBoxPeriods.Size = new System.Drawing.Size(414, 93);
            this.listBoxPeriods.TabIndex = 4;
            this.listBoxPeriods.SelectedValueChanged += new System.EventHandler(this.listBoxPeriods_SelectedValueChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(277, 39);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(48, 13);
            this.label5.TabIndex = 5;
            this.label5.Text = "Архивы:";
            // 
            // gridControl1
            // 
            this.gridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gridControl1.Location = new System.Drawing.Point(12, 179);
            this.gridControl1.MainView = this.gridView1;
            this.gridControl1.Name = "gridControl1";
            this.gridControl1.Size = new System.Drawing.Size(682, 160);
            this.gridControl1.TabIndex = 6;
            this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView1});
            // 
            // gridView1
            // 
            this.gridView1.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.PFMNameColumn,
            this.PartnerTypeColumn,
            this.FinPositionEPLColumn,
            this.SummColumn,
            this.DateStartColumn,
            this.DateEndColumn});
            this.gridView1.GridControl = this.gridControl1;
            this.gridView1.Name = "gridView1";
            this.gridView1.OptionsBehavior.AutoPopulateColumns = false;
            this.gridView1.OptionsView.ShowGroupPanel = false;
            // 
            // PFMNameColumn
            // 
            this.PFMNameColumn.Caption = "Структурное подразделение";
            this.PFMNameColumn.FieldName = "PFMName";
            this.PFMNameColumn.Name = "PFMNameColumn";
            this.PFMNameColumn.Visible = true;
            this.PFMNameColumn.VisibleIndex = 0;
            this.PFMNameColumn.Width = 165;
            // 
            // PartnerTypeColumn
            // 
            this.PartnerTypeColumn.Caption = "Тип контрагента";
            this.PartnerTypeColumn.FieldName = "PartnerType";
            this.PartnerTypeColumn.Name = "PartnerTypeColumn";
            this.PartnerTypeColumn.Visible = true;
            this.PartnerTypeColumn.VisibleIndex = 1;
            this.PartnerTypeColumn.Width = 122;
            // 
            // FinPositionEPLColumn
            // 
            this.FinPositionEPLColumn.Caption = "Код финансовой позиции ЕПЛ";
            this.FinPositionEPLColumn.FieldName = "FinPositionEPL";
            this.FinPositionEPLColumn.Name = "FinPositionEPLColumn";
            this.FinPositionEPLColumn.Visible = true;
            this.FinPositionEPLColumn.VisibleIndex = 2;
            this.FinPositionEPLColumn.Width = 77;
            // 
            // SummColumn
            // 
            this.SummColumn.Caption = "Сумма";
            this.SummColumn.FieldName = "Summ";
            this.SummColumn.Name = "SummColumn";
            this.SummColumn.Visible = true;
            this.SummColumn.VisibleIndex = 3;
            this.SummColumn.Width = 93;
            // 
            // DateStartColumn
            // 
            this.DateStartColumn.Caption = "Начало декады";
            this.DateStartColumn.FieldName = "DateStart";
            this.DateStartColumn.Name = "DateStartColumn";
            this.DateStartColumn.Visible = true;
            this.DateStartColumn.VisibleIndex = 4;
            this.DateStartColumn.Width = 99;
            // 
            // DateEndColumn
            // 
            this.DateEndColumn.Caption = "Конец декады";
            this.DateEndColumn.FieldName = "DateEnd";
            this.DateEndColumn.Name = "DateEndColumn";
            this.DateEndColumn.Visible = true;
            this.DateEndColumn.VisibleIndex = 5;
            this.DateEndColumn.Width = 105;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(19, 163);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(51, 13);
            this.label6.TabIndex = 7;
            this.label6.Text = "Декады:";
            // 
            // btn_ShowDetail
            // 
            this.btn_ShowDetail.Location = new System.Drawing.Point(280, 150);
            this.btn_ShowDetail.Name = "btn_ShowDetail";
            this.btn_ShowDetail.Size = new System.Drawing.Size(178, 23);
            this.btn_ShowDetail.TabIndex = 8;
            this.btn_ShowDetail.Text = "Детализация данных в Excel";
            this.btn_ShowDetail.UseVisualStyleBackColor = true;
            this.btn_ShowDetail.Click += new System.EventHandler(this.btn_ShowDetail_Click);
            // 
            // ReportDate
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(706, 348);
            this.Controls.Add(this.btn_ShowDetail);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.gridControl1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.listBoxPeriods);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btn_OK);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dateTimePicker2);
            this.Controls.Add(this.dateTimePicker1);
            this.Name = "ReportDate";
            this.Text = "Выбор периода ";
            this.Load += new System.EventHandler(this.ReportDate_Load);
            ((System.ComponentModel.ISupportInitialize)(this.listBoxPeriods)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btn_OK;
        private System.Windows.Forms.Label label4;
        private DevExpress.XtraEditors.ListBoxControl listBoxPeriods;
        private System.Windows.Forms.Label label5;
        private DevExpress.XtraGrid.GridControl gridControl1;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
        private System.Windows.Forms.Label label6;
        private DevExpress.XtraGrid.Columns.GridColumn PFMNameColumn;
        private DevExpress.XtraGrid.Columns.GridColumn PartnerTypeColumn;
        private DevExpress.XtraGrid.Columns.GridColumn FinPositionEPLColumn;
        private DevExpress.XtraGrid.Columns.GridColumn SummColumn;
        private DevExpress.XtraGrid.Columns.GridColumn DateStartColumn;
        private DevExpress.XtraGrid.Columns.GridColumn DateEndColumn;
        private System.Windows.Forms.Button btn_ShowDetail;
    }
}