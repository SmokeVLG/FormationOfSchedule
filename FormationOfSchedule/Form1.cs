using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Collections;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using System.Data.Common;
//using System.Windows.Forms;
//using System.Configuration;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.Drawing;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.Utils;
using FormationOfSchedule.Properties;
using System.Data.Odbc;

namespace FormationOfSchedule
{
    public partial class Form1 : Form
    {
        string login_;
        string password_;
        string TemplatePath;
        Excel.Application app = null;       // представляет само приложение Excel
        Excel.Workbook theBook;             // позволяет работать со всеми открытыми рабочими книгами, создавать рабочую книгу и импортировать данные в новую рабочую книгу
        //private SqlConnection _TestCon;
        private SqlConnection _connection;
        Form temp; 
        public Form1(string login, string password, Form main)
        {
            InitializeComponent();
          //  _command = new OleDbCommand();
            if (_connection == null)
            {
                temp = main;
                try
                {
                    _connection = new SqlConnection(ConfigurationManager.ConnectionStrings["FormationOfSchedule.Properties.Settings.FormationOfSchedule"].ConnectionString);
                    SqlConnectionStringBuilder scsb = new SqlConnectionStringBuilder();
                    scsb.DataSource = _connection.DataSource;
                    scsb.InitialCatalog = _connection.Database;
                    scsb.UserID = login;
                    scsb.Password = password;
                    login_ = login;
                    password_ = password;
                    _connection = new SqlConnection(scsb.ConnectionString);
                    string[] inform = new string[3];
                    DataSaver ds = new DataSaver(login_, password_);
                    inform = ds.getPFM();
                    inform_ = inform;

                    if (inform_[2] == "1")
                    {
                        blockMark = "unlock";
                        btn_block.Text = "Разблокировать ввод ПП";
                    }
                    else
                    {
                        blockMark = "lock";
                        btn_block.Text = "Заблокировать ввод ПП";
                    }
                    
					//if (inform_[0].Substring(0, 13) != "Администратор") Изменил Дороненков Г.Г. 03-06-2014 
					if ((inform_[0] != string.Empty) && (inform_[0].Substring(0, 13) != "Администратор"))
                    {
                        //tabPage2.Enabled = false;
                        btn_del_actual.Visible = false;
                        lb_limits_file.Visible = false;
                        btn_limits_load.Visible = false;
                        groupBox3.Visible = false;
                        groupBox7.Visible = false;
                        btnAdd.Visible = false;
                        linkAdd.Visible = false;
                        linkDel.Visible = false;
                        linkChenge.Visible = false;
                       // limitsControl.Location.Y = System.Drawing.Point.
                       // limitsControl.Location = new Point(6, 10);
                       // tabPage4.Enabled = false;
                        btn_actual_load.Visible = false;
                        lb_actual_file.Visible = false;
                        gb_actual.Visible = false;
                      //  ActualControl.Location = new Point(6, 10);
                        tabPage5.Enabled = false;

                        btn_block.Visible = false;
                        otch2.Visible = false;
                        Otch3.Visible = false;
                        otch_sv.Visible = false;
                        otch5.Visible = false;
                        otch6.Visible = false;

                        dateTimePicker10.Visible = false;
                        label68.Visible = false;
                        dateTimePicker7.Visible = false;
                        label58.Visible = false;

                        if (inform_[2] == "1")
                        {
                            button7.Enabled = false;
                           // MessageBox.Show("Ввод планируемых платежей заблокирован!", "Внимание!!!");
                            lBlockMessage.Visible = true;
                        }
                    }
                    tb_pay_ContragentType.ReadOnly = true;
                    tb_pay_EPL.ReadOnly = true;
                    tb_pay_PFMcode.ReadOnly = true;
                    tb_pay_SummRUB.ReadOnly = true;
                    tb_plan_statelikvid.ReadOnly = true;
                    tb_plan_stavrolen.ReadOnly = true;
                    tb_pay_ContragentType.BackColor = System.Drawing.Color.White;
                    tb_pay_EPL.BackColor = System.Drawing.Color.White;
                    tb_pay_PFMcode.BackColor = System.Drawing.Color.White;
                    tb_pay_SummRUB.BackColor = System.Drawing.Color.White;
                    tb_plan_stavrolen.BackColor = System.Drawing.Color.White;
                    tb_plan_statelikvid.BackColor = System.Drawing.Color.White;
                    tb_pay_PFMcode.Text = inform[1];
                    TemplatePath = Directory.GetCurrentDirectory() + "\\Reports\\";

                    string separator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
                    if (separator == ".")
                    {
                        separator_true = '.';
                        separator_false = ',';
                    }
                    else
                    {
                        separator_true = ',';
                        separator_false = '.';
                    }
                }
                catch (Exception ex)
                {
                  //  MessageBox.Show("Проблема с авторизацией! Возможно у Вас нет доступа к ПО.");

                       MessageBox.Show(ex.Message);
                    StringError = "error";
                }
            }
            
            

        }
        char separator_false;
        char separator_true;
        string StringError;
        string[] inform_ = new string[3];
        string link = "INS";
        string operation = "INS";
        private string _fileName;
        //private SqlConnection _opRegCon;
        /// <summary>Запрос
        /// </summary>
       // private OleDbCommand _command;
        //private SqlCommand _command;
       // private OleDbDataReader _reader;
       // private SqlCommand _reader;
        List<string> _sheets = new List<string>();
        private List<Data> _data = new List<Data>();
       // private List<Data> _toDelete = new List<Data>();
       // private List<Data> _toProtocol = new List<Data>();
       // private OleDbConnection _fileConn;
        /// <summary>Имя файла с планами
        /// </summary>
        public string FileName
        {
            get
            {
                return _fileName;
            }
            set
            {
                _fileName = value;
            }
        }

        string _sheet;
        string _range;
//private  Color PaleTurquoise; 

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpFilD = new OpenFileDialog();
            if (OpFilD.ShowDialog() == DialogResult.OK)
            {
                label1.Text = OpFilD.FileName;
                _fileName = OpFilD.FileName;
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += new DoWorkEventHandler(bw_DoWork);
                bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);
                groupBox1.Enabled = false;
                Cursor = Cursors.WaitCursor;
                bw.RunWorkerAsync();
            }

        }

            void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
            this.listBox1.DataSource = null;
            this.listBox1.DataSource = _sheets;
            this.groupBox1.Enabled = true;
            Cursor = Cursors.Arrow;
        }

        void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                theBook = app.Workbooks.Open(_fileName, Type.Missing,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка при открытии книги");
            }

            try
            {
                _sheets.Clear();
                foreach (Excel.Worksheet w in theBook.Worksheets)
                {
                    _sheets.Add(w.Name);
                }
                // Закрываем Excel
                theBook.Close(false, Type.Missing, Type.Missing);
                app.Quit();
                theBook = null;
                app = null;
                GC.Collect();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
            }

        }


/*
        private void FullGrid(string str)
        {

            ArrayList listRow = new ArrayList();


            int i = 0;
            string[] pice = new string[0];
            listRow.AddRange(str.Split(';'));

            MessageBox.Show("listRow[0] = " + listRow[1]);
        }
        */
        private void btn_Close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_Insert_Click(object sender, EventArgs e)
        {
            if (listBox1.Text != "")
            {
                int rows;
                if (int.TryParse(textBox2.Text, out rows))
                {
                    //начать загрузку
                    if (textBox1.Text == "")
                        textBox1.Text = "1";
                    _range = "A" + (Convert.ToInt32(textBox1.Text) - 1).ToString() + ":V" + textBox2.Text;
                    OdbcConnection cn = new OdbcConnection();
                    cn.ConnectionString = string.Format(@"Driver={{Microsoft Excel Driver (*.xls)}};DBQ={0};ReadOnly=0;", label1.Text);
                    string strCom = "select * from [" + _sheet + "$" + _range + "]";
                    cn.Open();
                    OdbcCommand comm_mon = new OdbcCommand(strCom, cn);
                    OdbcDataAdapter da = new OdbcDataAdapter();
                    da.SelectCommand = comm_mon;
                    System.Data.DataTable dt = new System.Data.DataTable();
                    da.Fill(dt);
                    DataSaver ds = new DataSaver(login_, password_);
                    string Warning = "";
                    Warning = ds.Save(dt, Convert.ToInt32(textBox1.Text));
                    ContractsFill();
                    if (Warning != "")
                        MessageBox.Show(Warning);
                }
                else
                {
                    MessageBox.Show("Укажите количество строк");
                }
            }
        }

        private void ContractsFill()
        {
            string date = dateTimePicker10.Value.Date.Year.ToString() + "-" + dateTimePicker10.Value.Date.Month.ToString() + "-" + dateTimePicker10.Value.Date.Day.ToString();
            string Comm = "";
            //if (inform_[0].Substring(0, 13) == "Администратор") Изменил Дороненков Г.Г. 03-06-2014 
			if ((inform_[0] != string.Empty) && (inform_[0].Substring(0, 13) == "Администратор"))
            {
                Comm = "SELECT num ,ID  , col1 ,col2 ,col3,col4 ,col19 ,col5,col6,col7,col8,col9,col10,col11,col12,col13,col14,col15,col16,col17,col18,col20  FROM [udf_FS_Contracts_Get] ()" +
                             "  WHERE col2 not in (select ContractCode from Contracts where ContractStatus not in('В работе','Корректировка','На проверке','На экспертизе','Проект','Исправление')) " + 
                             " AND '" + date + "' between col12 and col14";
            }
            else
            {
                Comm = "SELECT num ,ID  , col1 ,col2 ,col3,col4 ,col19 ,col5,col6,col7,col8,col9,col10,col11,col12,col13,col14,col15,col16,col17,col18,col20  FROM [udf_FS_Contracts_Get] ()" +
                 "   WHERE col2 not in (select ContractCode from Contracts where ContractStatus not in('В работе','Корректировка','На проверке','На экспертизе','Проект','Исправление')) and col7 = '" + inform_[1] + "'";
            }
            SqlCommand comm = new SqlCommand(Comm, _connection);
            try
            {
                if (_connection.State != System.Data.ConnectionState.Open)
                    _connection.Open();
                comm.ExecuteScalar();

                SqlDataAdapter new_dataAdapter = new SqlDataAdapter(comm);
                var dataSet = new DataSet();
                new_dataAdapter.Fill(dataSet, "contractsFill");
                ContractControl.DataSource = dataSet;
                gridControl1.DataSource = dataSet;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            _connection.Close();
        }


        private void LimitsFill()
        {
            string monthList = "(1)";

            if (comboBox1.Text == "1 квартал")
                monthList = "(1, 2, 3)";
            else if (comboBox1.Text == "2 квартал")
                monthList = "(4, 5, 6)";
            else if (comboBox1.Text == "3 квартал")
                monthList = "(7, 8, 9)";
            else if (comboBox1.Text == "4 квартал")
                monthList = "(10, 11, 12)";
            else monthList = "(1)";

            string Comm = "";
            
			//if (inform_[0].Substring(0, 13) == "Администратор") Изменил Дороненков Г.Г. 03-06-2014 
			if ((inform_[0] != string.Empty) && (inform_[0].Substring(0, 13) == "Администратор"))
            {
                 Comm = "SELECT l.[IdLimits] as limID ,l.[PFMcode] as lim1 ,l.[FinPositionEPL] as lim2 ,l.[Year] as lim3" +
                                           " ,m.[Month] as lim4 ," +
                                           "(select chislo from udf_Convert_Float_to_Char(l.Summ)) as lim5" +
                                           ",l.[CurrencyRUB] as lim6, l.[Month] as limMonth FROM [Limits] l" +
                                           " LEFT JOIN Month m on l.Month = m.numMonth  WHERE l.Month in " + monthList + " and l.Year = '" + textBox3.Text + "'";
            }
            else
            {
               Comm = "SELECT l.[IdLimits] as limID ,l.[PFMcode] as lim1 ,l.[FinPositionEPL] as lim2 ,l.[Year] as lim3" +
                                        " ,m.[Month] as lim4 ," +
                                        "(select chislo from udf_Convert_Float_to_Char(l.Summ)) as lim5 " + 
                                        ",l.[CurrencyRUB] as lim6, l.[Month] as limMonth FROM [Limits] l" +
                                        " LEFT JOIN Month m on l.Month = m.numMonth WHERE PFMcode = '" + inform_[1] + "' and l.Month in " + monthList + " and l.Year = '" + textBox3.Text + "'";
            }
            SqlCommand limitsComm = new SqlCommand(Comm, _connection);
                if (_connection.State != System.Data.ConnectionState.Open)
                    _connection.Open();
                limitsComm.ExecuteScalar();

                SqlDataAdapter new_dataAdapter = new SqlDataAdapter(limitsComm);
                var dataSet = new DataSet();
                new_dataAdapter.Fill(dataSet, "limitsFill");
                limitsControl.DataSource = dataSet;
           // limitsView.Columns.  Column("lim5").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
           // limitsView.Columns["lim5"].CellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
           // limitsView.Columns["lim5"].Name = DataGridViewContentAlignment.BottomCenter.ToString();
           // this.limitsView.Columns["lim5"]
              DataGridViewCellStyle renk = new DataGridViewCellStyle();
              renk.Alignment = DataGridViewContentAlignment.BottomRight;


         /*
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            */
            _connection.Close();
        }

        private void ActualPaymentsFill()
        {

                string date1 = dateTimePicker8.Value.Date.Year.ToString() + "-" + dateTimePicker8.Value.Date.Month.ToString() + "-" + dateTimePicker8.Value.Date.Day.ToString();
                string date2 = dateTimePicker9.Value.Date.Year.ToString() + "-" + dateTimePicker9.Value.Date.Month.ToString() + "-" + dateTimePicker9.Value.Date.Day.ToString();
                string comm = "";
                
			    //if (inform_[0].Substring(0, 13) == "Администратор") Изменил Дороненков Г.Г. 03-06-2014 
				if ((inform_[0] != string.Empty) && (inform_[0].Substring(0, 13) == "Администратор"))
                {
                    comm = "SELECT [Idactual] ,[ContractCode] ,[PartnerCode] ,[PartnerType],[FinPosition],FPStavrolenName ,[FinPositionEPL],StateEverydayLicvid,[PFMcode] ,[DatePayments],[Currency] ,"
                            + "(select chislo from udf_Convert_Float_to_Char([Summ])) as [Summ] " +
                            ",(select chislo from udf_Convert_Float_to_Char([SummRus])) as [SummRus] " +
                                                        "FROM [ActualPayments] LEFT JOIN FinancialPosition ON FinancialPosition.FPcode = FinPositionEPL AND FinancialPosition.FPcodeStavrolen =  FinPosition " +
                                                        " WHERE ActualPayments.DatePayments between '" + date1 + "' and '" + date2 + "'";
                }

                else
                {
                    comm = "SELECT [Idactual] ,[ContractCode] ,[PartnerCode] ,[PartnerType],[FinPosition],FPStavrolenName ,[FinPositionEPL],StateEverydayLicvid,[PFMcode] ,[DatePayments],[Currency] " +
                            ",(select chislo from udf_Convert_Float_to_Char([Summ])) as [Summ] " +
                            ",(select chislo from udf_Convert_Float_to_Char([SummRus])) as [SummRus]  " +
                                            "FROM [ActualPayments] LEFT JOIN FinancialPosition ON FinancialPosition.FPcode = FinPositionEPL AND FinancialPosition.FPcodeStavrolen =  FinPosition " +
                                            "WHERE PFMcode = '" + inform_[1] + "' and " +
                                            " ActualPayments.DatePayments between '" + date1 + "' and '" + date2 + "'";
                }

                SqlCommand Comm = new SqlCommand(comm, _connection);
                try
                {
                    if (_connection.State != System.Data.ConnectionState.Open)
                        _connection.Open();
                    Comm.ExecuteScalar();

                    SqlDataAdapter new_dataAdapter = new SqlDataAdapter(Comm);
                    var dataSet = new DataSet();
                    new_dataAdapter.Fill(dataSet, "ActualFill");
                    ActualControl.DataSource = dataSet;

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Проверьте правильность периода выгрузки данных! " +ex.Message);
                }

                _connection.Close();
            
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.listBox1.DataSource != null)
            {
                try
                {
                    _sheet = this.listBox1.SelectedValue.ToString();
                    //_sheet = _sheet.Substring(0, _sheet.Length - 2);
                    char[] charsToTrim = { '$', '#' };
                    _sheet = _sheet.Trim(charsToTrim);
                }
                catch { }
            }
        }

        public void timer1_Tick(object sender, EventArgs e)
        {
            lb_status.Text = "";
            DataSaver ds = new DataSaver(login_, password_);
            inform_ = ds.getPFM();
			//if (inform_[2] == "1" && (inform_[0].Substring(0, 13) != "Администратор"))  Изменил Дороненков Г.Г. 03-06-2014 
			if (inform_[2] == "1" && ((inform_[0] != string.Empty) && (inform_[0].Substring(0, 13) != "Администратор")))
            {
                button7.Enabled = false;
                lBlockMessage.Visible = true;
            }
            else
            {
                button7.Enabled = true;
                lBlockMessage.Visible = false;
            }

           // MessageBox.Show(DateTime.Now.ToString());

            /*
            if (DateTime.Now.ToString("HH:mm") == "13:05" || DateTime.Now.ToString("HH:mm") == "17:05")
            {
                _connection.Close();
                timer1.Stop();
                temp.Close();
            }
            */
        }

        private void InstallTime()
        {
            int day;
            int month;
            int year;

            if (DateTime.Today.Month == 1)
            {
                day = DateTime.Today.Day;
                month = 12;
                year = DateTime.Today.Year - 1;
            }
            else
            {
                day = DateTime.Today.Day;
                month = DateTime.Today.Month - 1;
                if (day == 28 || day == 29 || day == 30 || day == 31)
                {
                    day = DateTime.Today.Day - 3;
                }
                year = DateTime.Today.Year;
            }
            dateTimePicker8.Value = new DateTime(year, month, day);

            if (DateTime.Today.Month == 12)
            {
                day = DateTime.Today.Day;
                month = 1;
                year = DateTime.Today.Year + 1;
            }
            else
            {
                day = DateTime.Today.Day;
                month = DateTime.Today.Month + 1;
                if (day == 28 || day == 29 || day == 30 || day == 31)
                {
                    day = DateTime.Today.Day - 3;
                }
                year = DateTime.Today.Year;
            }
            dateTimePicker9.Value = new DateTime(year, month, day);

            if (DateTime.Today.Month == 11)
            {
                day = DateTime.Today.Day;
                month = 1;
                year = DateTime.Today.Year + 1;
            }
            else if (DateTime.Today.Month == 12)
            {
                day = DateTime.Today.Day;
                month = 2;
                if (day == 28 || day == 29 || day == 30 || day == 31)
                {
                    day = DateTime.Today.Day - 3;
                }
                year = DateTime.Today.Year + 1;
            }
            else
            {
                day = DateTime.Today.Day;
                month = DateTime.Today.Month + 2;
                if (day == 28 || day == 29 || day == 30 || day == 31)
                {
                    day = DateTime.Today.Day - 3;
                }
                year = DateTime.Today.Year;
            }
            dateTimePicker1.Value = new DateTime(year, month, day);

            if (DateTime.Today.Month == 11)
            {
                day = DateTime.Today.Day;
                month = 1;
                year = DateTime.Today.Year + 1;
            }
            else if (DateTime.Today.Month == 12)
            {
                day = DateTime.Today.Day;
                month = 2;
                if (day == 28 || day == 29 || day == 30 || day == 31)
                {
                    day = DateTime.Today.Day - 3;
                }
                year = DateTime.Today.Year + 1;
            }
            else
            {
                day = DateTime.Today.Day;
                month = DateTime.Today.Month + 2;
                if (day == 28 || day == 29 || day == 30 || day == 31)
                {
                    day = DateTime.Today.Day - 3;
                }
                year = DateTime.Today.Year;
            }
            dateTimePicker1.Value = new DateTime(year, month, day);

            textBox3.Text = DateTime.Today.Year.ToString();

            if(DateTime.Today.Month == 1 || DateTime.Today.Month == 2 || DateTime.Today.Month == 3)
                comboBox1.SelectedItem = "1 квартал";
            else if(DateTime.Today.Month == 4 || DateTime.Today.Month == 5 || DateTime.Today.Month == 6)
                comboBox1.SelectedItem = "2 квартал";
            else if (DateTime.Today.Month == 7 || DateTime.Today.Month == 8 || DateTime.Today.Month == 9)
                comboBox1.SelectedItem = "3 квартал";
            else if (DateTime.Today.Month == 10 || DateTime.Today.Month == 11 || DateTime.Today.Month == 12)
                comboBox1.SelectedItem = "4 квартал";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'formOfShDataSet1.Last4Date' table. You can move, or remove it, as needed.
           // this.last4DateTableAdapter1.Fill(this.formOfShDataSet1.Last4Date);
            // TODO: This line of code loads data into the 'formOfShDataSet.Last4Date' table. You can move, or remove it, as needed.
           // this.last4DateTableAdapter.Fill(this.formOfShDataSet.Last4Date);
            if (StringError == "error")
            { temp.Close(); }
            else
            {
                try
                {
                    InstallTime();
                }
                catch { }
                // TODO: This line of code loads data into the 'formationOfScheduleDataSet6.v_FinancialPosition_Topic' table. You can move, or remove it, as needed.
                this.v_FinancialPosition_TopicTableAdapter.Fill(this.formationOfScheduleDataSet6.v_FinancialPosition_Topic);
                // TODO: This line of code loads data into the 'formationOfScheduleDataSet5.v_FinPosition_EPL' table. You can move, or remove it, as needed.
                this.v_FinPosition_EPLTableAdapter.Fill(this.formationOfScheduleDataSet5.v_FinPosition_EPL);
                // TODO: This line of code loads data into the 'udf_getcurrency.udf_FS_get_currency' table. You can move, or remove it, as needed.
                this.udf_FS_get_currencyTableAdapter.Fill(this.udf_getcurrency.udf_FS_get_currency);
                // TODO: This line of code loads data into the 'formationOfScheduleDataSet4.v_FPcodeStavrolen' table. You can move, or remove it, as needed.
                this.v_FPcodeStavrolenTableAdapter.Fill(this.formationOfScheduleDataSet4.v_FPcodeStavrolen);
                // TODO: This line of code loads data into the 'formationOfScheduleDataSet3.FinancialPosition' table. You can move, or remove it, as needed.
                this.financialPositionTableAdapter.Fill(this.formationOfScheduleDataSet3.FinancialPosition);
                // TODO: This line of code loads data into the 'currencyCurs1.v_CurrencyCurs' table. You can move, or remove it, as needed.
                this.v_CurrencyCursTableAdapter.Fill(this.currencyCurs1.v_CurrencyCurs);
                // TODO: This line of code loads data into the 'formationOfScheduleDataSet2.v_Currency' table. You can move, or remove it, as needed.
                this.v_CurrencyTableAdapter.Fill(this.formationOfScheduleDataSet2.v_Currency);
                // TODO: This line of code loads data into the 'currency_.CurrencyCurs' table. You can move, or remove it, as needed.
                this.currencyCursTableAdapter1.Fill(this.currency_.CurrencyCurs);
                // TODO: This line of code loads data into the 'currencyCurs._CurrencyCurs' table. You can move, or remove it, as needed.
                this.currencyCursTableAdapter.Fill(this.currencyCurs._CurrencyCurs);
                // TODO: This line of code loads data into the 'fS_UsersGroup.UsersGroup' table. You can move, or remove it, as needed.
                this.usersGroupTableAdapter.Fill(this.fS_UsersGroup.UsersGroup);
                // TODO: This line of code loads data into the 'formationOfScheduleDataSet1.v_Users_GET' table. You can move, or remove it, as needed.
                this.v_Users_GETTableAdapter.Fill(this.formationOfScheduleDataSet1.v_Users_GET);
                // TODO: This line of code loads data into the 'formationOfScheduleDataSet.Month' table. You can move, or remove it, as needed.
                this.monthTableAdapter.Fill(this.formationOfScheduleDataSet.Month);
                ContractsFill();
                LimitsFill();
                PlansFill();
                ActualPaymentsFill();
                KSSStoContragentFill();
                CurrencyView.Columns[0].SortOrder = DevExpress.Data.ColumnSortOrder.Descending;
                timer1.Start();
            }
        }


        private void PlansFill()
        {
                string date1 = dateTimePicker3.Value.Date.Year.ToString() + "-" + dateTimePicker3.Value.Date.Month.ToString() + "-" + dateTimePicker3.Value.Date.Day.ToString();
                string date2 = dateTimePicker1.Value.Date.Year.ToString() + "-" + dateTimePicker1.Value.Date.Month.ToString() + "-" + dateTimePicker1.Value.Date.Day.ToString();
                string Comm = "";

				//if (inform_[0].Substring(0, 13) == "Администратор") Изменил Дороненков Г.Г. 03-06-2014 
				if ((inform_[0] != string.Empty) && (inform_[0].Substring(0, 13) == "Администратор"))

                {
                    Comm = " Select [IdPay] as planID ,[ContractCode] as plan1 ,[PartnerCode] as plan2" +
                                ",[PartnerType] as plan3 ,[FinPosition] as plan4 ,[FinPositionEPL] as plan5" +
                                ",[PFMcode] as plan6 ,[DatePay] as plan7 ,[PayCurrency] as plan8" +
                                ",(select chislo from udf_Convert_Float_to_Char([PaySumm])) as plan9" +
                                ",(select chislo from udf_Convert_Float_to_Char([PaySummRus])) as plan10" +
                                ", FinancialPosition.StateEverydayLicvid as plan11" +
                                ",FinancialPosition.FPStavrolenName as plan12, Comment as planComment  FROM PaymentsPlan" +
                                " LEFT JOIN FinancialPosition ON FinancialPosition.FPcode = PaymentsPlan.FinPositionEPL" +
                                " and FinancialPosition.FPcodeStavrolen = PaymentsPlan.FinPosition" +
                                " WHERE DatePay between '" + date1 + "' and '" + date2 + "'";
                }

                else
                {
                    Comm = " Select [IdPay] as planID ,[ContractCode] as plan1 ,[PartnerCode] as plan2" +
                                ",[PartnerType] as plan3 ,[FinPosition] as plan4 ,[FinPositionEPL] as plan5" +
                                ",[PFMcode] as plan6 ,[DatePay] as plan7 ,[PayCurrency] as plan8" +
                                ",(select chislo from udf_Convert_Float_to_Char([PaySumm])) as plan9" +
                                ",(select chislo from udf_Convert_Float_to_Char([PaySummRus])) as plan10 " +
                                ", FinancialPosition.StateEverydayLicvid as plan11" +
                                ",FinancialPosition.FPStavrolenName as plan12, Comment as planComment  FROM PaymentsPlan" +
                                " LEFT JOIN FinancialPosition ON FinancialPosition.FPcode = PaymentsPlan.FinPositionEPL" +
                                " and FinancialPosition.FPcodeStavrolen = PaymentsPlan.FinPosition" +
                                " WHERE PFMcode = '" + inform_[1] + "' " +
                                " and DatePay between '" + date1 + "' and '" + date2 + "'";
                }
                SqlCommand PlansComm = new SqlCommand(Comm, _connection);
                try
                {
                    if (_connection.State != System.Data.ConnectionState.Open)
                        _connection.Open();
                    PlansComm.ExecuteScalar();

                    SqlDataAdapter new_dataAdapter = new SqlDataAdapter(PlansComm);
                    var dataSet = new DataSet();

                    new_dataAdapter.Fill(dataSet, "PlansFill");
                    gridControl2.DataSource = dataSet;

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Проверьте правильность периода выгрузки данных! " + ex.Message);
                }

                _connection.Close();
                gridView2.FocusedRowHandle = focused_PaymentsPlan_gridView2;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            LimitsFill();
            tb_finpos.Text = "";
            tb_PFM.Text = "";
            tb_summ.Text = "";
            tb_year.Text = "";
            lb_status.Text = ""; 
        }

       /*
        private void limitsControl_DataSourceChanged(object sender, EventArgs e)
        {
           // MessageBox.Show("Save change?");
        }*/
        /*
        private void limitsView_CellValueChanged(object sender, CellValueChangedEventArgs e)
        {
            MessageBox.Show("Save change?");

        }*/

        private void button2_Click(object sender, EventArgs e)
        {
            _connection.Close();
            temp.Close();
            //this.Close();
            timer1.Stop();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            int focused_row = 0;
            if (link == "INS")
            {

                if (tb_PFM.Text == "" || tb_finpos.Text == "" || tb_summ.Text == "")
                {
                    lb_status.Text = "Ошибка: Заполните поля Код ПФМ, Финансовая позиция и Сумма!";
                }

                else
                {
                    DataSaver ds = new DataSaver(login_, password_);
                    lb_status.Text = ds.AddLimits
                                                (tb_PFM.Text.Replace(" ", "")
                                                , tb_finpos.Text.Replace(" ", "")
                                                , tb_year.Text.Replace(" ", "")
                                                 , cb_month.SelectedValue.ToString()
                                                 , tb_summ.Text.Replace(separator_false, separator_true).Replace(" ", "")
                                                );
                }

                tb_PFM.Text = "";
                tb_year.Text = "";
                tb_finpos.Text = "";
                tb_summ.Text = "";
                focused_row = limitsView.RowCount;
            }

            if (link == "UPD")
            {
                focused_row = limitsView.FocusedRowHandle;
                if (tb_PFM.Text == "" || tb_finpos.Text == "" || tb_summ.Text == "")
                {
                    lb_status.Text = "Ошибка: Заполните поля Код ПФМ, Финансовая позиция и Сумма!";
                }

                else
                {
                    DataSaver ds = new DataSaver(login_, password_);
                    lb_status.Text = ds.UpdateLimits
                                                (tb_PFM.Text.Replace(" ", "")
                                                , tb_finpos.Text.Replace(" ", "")
                                                , tb_year.Text.Replace(" ", "")
                                                 , cb_month.SelectedValue.ToString()
                                                 , tb_summ.Text.Replace(separator_false, separator_true).Replace(" ", "")
                                                 ,tb_id.Text
                                                );
                }

            }

            if (link == "DEL")
            {
                try
                {
                    if (limitsView.FocusedRowHandle != limitsView.RowCount - 1)
                        focused_row = limitsView.FocusedRowHandle;
                    else
                        focused_row = limitsView.FocusedRowHandle - 1;
                }
                catch { }
                string idRow = "";
                idRow = limitsView.GetFocusedDataRow()["limID"].ToString();
                DataSaver ds = new DataSaver(login_, password_);
                lb_status.Text = ds.DeleteLimits(idRow);

                _connection.Close();
            }

            LimitsFill();
            limitsView.FocusedRowHandle = focused_row;
        }

        private void gridContextMenuHandler(object sender, EventArgs e)
        {
            string idRow;

            if (((ContextMenuStrip)((ToolStripMenuItem)sender).Owner).SourceControl.Equals(limitsControl))
            {
          //  switch (((ToolStripMenuItem)sender).Name)
           // {
             //   case "DelRowMenuItem":
               //     {
                idRow = limitsView.GetFocusedDataRow()["limID"].ToString();
                        SqlCommand comm = new SqlCommand("usp_FS_Del_in_Limits", _connection);
                        comm.CommandType = System.Data.CommandType.StoredProcedure;

                        SqlParameter tParam = new SqlParameter("@id", SqlDbType.Int);
                        tParam.Value = idRow;
                        comm.Parameters.Add(tParam);
                        
                        tParam = new SqlParameter("@res", SqlDbType.Char);
                        tParam.Direction = ParameterDirection.Output;
                        tParam.Value = "";
                        tParam.Size = 50;
                        comm.Parameters.Add(tParam);

                        if (_connection.State != System.Data.ConnectionState.Open)
                            _connection.Open();
                        comm.ExecuteNonQuery();

                        lb_status.Text = comm.Parameters["@res"].Value.ToString();

                        _connection.Close();

                        LimitsFill();

                     //  break;
                  //  }

            }
        }

        private void linkChenge_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (limitsView.RowCount != 0)
            {
                link = "UPD";
                linkAdd.LinkVisited = false;
                linkChenge.LinkVisited = true;
                linkDel.LinkVisited = false;
                btnAdd.Text = "Изменить";
                groupBox3.Enabled = true;
                tb_PFM.Text = limitsView.GetFocusedDataRow()["lim1"].ToString();
                tb_finpos.Text = limitsView.GetFocusedDataRow()["lim2"].ToString();
                tb_year.Text = limitsView.GetFocusedDataRow()["lim3"].ToString();
                tb_summ.Text = limitsView.GetFocusedDataRow()["lim5"].ToString();
                cb_month.SelectedValue = limitsView.GetFocusedDataRow()["limMonth"].ToString();
                tb_id.Text = limitsView.GetFocusedDataRow()["limID"].ToString();
            }
        }
        int focused_limits = 1;
        private void limitsView_FocusedRowChanged(object sender, FocusedRowChangedEventArgs e)
        {
            focused_limits = limitsView.FocusedRowHandle;
            if (link == "UPD" || link == "DEL")
            {
                try
                {
                    tb_PFM.Text = limitsView.GetFocusedDataRow()["lim1"].ToString();
                    tb_finpos.Text = limitsView.GetFocusedDataRow()["lim2"].ToString();
                    tb_year.Text = limitsView.GetFocusedDataRow()["lim3"].ToString();
                    tb_summ.Text = limitsView.GetFocusedDataRow()["lim5"].ToString();
                    cb_month.SelectedValue = limitsView.GetFocusedDataRow()["limMonth"].ToString();
                    tb_id.Text = limitsView.GetFocusedDataRow()["limID"].ToString();
                }
                catch { }
            }
        }

        private void linkAdd_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link = "INS";
            linkAdd.LinkVisited = true;
            linkChenge.LinkVisited = false;
            btnAdd.Text = "Добавить";
            groupBox3.Enabled = true;
            linkDel.LinkVisited = false;
        }

        private void linkUsAdd_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            operation = "INS";
            btn_Users.Text = "Добавить";
            linkUsAdd.LinkVisited = true;
            linkUsUpdate.LinkVisited = false;
            linkUsDel.LinkVisited = false;
            grB_Users.Enabled = true;
        }

        int focused_users = 1;
        private void usersView_FocusedRowChanged_1(object sender, FocusedRowChangedEventArgs e)
        {
            focused_users = usersView.FocusedRowHandle;

            try
            {
                if (operation == "UPD" || operation == "DEL")
                {
                    tb_login.Text = usersView.GetFocusedDataRow()["Login"].ToString();
                    tb_PFM_users.Text = usersView.GetFocusedDataRow()["PFMcode"].ToString();
                    tb_PFMname.Text = usersView.GetFocusedDataRow()["PFMname"].ToString();
                    tb_fio.Text = usersView.GetFocusedDataRow()["UsersName"].ToString();
                    cb_groupUsers.SelectedValue = usersView.GetFocusedDataRow()["Expr1"].ToString();
                    tb_users_ID.Text = usersView.GetFocusedDataRow()["IdUsers"].ToString();
                }
            }
            catch { }
        }

        private void btn_users_New_Click(object sender, EventArgs e)
        {
            this.v_Users_GETTableAdapter.Fill(this.formationOfScheduleDataSet1.v_Users_GET);
            tb_login.Text = "";
            tb_PFM_users.Text = "";
            tb_PFMname.Text = "";
            tb_fio.Text = "";
            tb_users_ID.Text = "";
        }

        private void linkUsUpdate_LinkClicked(object sender, EventArgs e)
        {
            if (usersView.RowCount != 0)
            {
                operation = "UPD";
                btn_Users.Text = "Изменить";
                linkUsAdd.LinkVisited = false;
                linkUsUpdate.LinkVisited = true;
                linkUsDel.LinkVisited = false;

                tb_login.Text = usersView.GetFocusedDataRow()["Login"].ToString();
                tb_PFM_users.Text = usersView.GetFocusedDataRow()["PFMcode"].ToString();
                tb_PFMname.Text = usersView.GetFocusedDataRow()["PFMname"].ToString();
                tb_fio.Text = usersView.GetFocusedDataRow()["UsersName"].ToString();
                cb_groupUsers.SelectedValue = usersView.GetFocusedDataRow()["Expr1"].ToString();
                tb_users_ID.Text = usersView.GetFocusedDataRow()["IdUsers"].ToString();

                grB_Users.Enabled = true;
            }
        }

        private void linkUsDel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (usersView.RowCount != 0)
            {
                operation = "DEL";
                btn_Users.Text = "Удалить";
                linkUsAdd.LinkVisited = false;
                linkUsUpdate.LinkVisited = false;
                linkUsDel.LinkVisited = true;

                tb_login.Text = usersView.GetFocusedDataRow()["Login"].ToString();
                tb_PFM_users.Text = usersView.GetFocusedDataRow()["PFMcode"].ToString();
                tb_PFMname.Text = usersView.GetFocusedDataRow()["PFMname"].ToString();
                tb_fio.Text = usersView.GetFocusedDataRow()["UsersName"].ToString();
                cb_groupUsers.SelectedValue = usersView.GetFocusedDataRow()["Expr1"].ToString();
                tb_users_ID.Text = usersView.GetFocusedDataRow()["IdUsers"].ToString();

                grB_Users.Enabled = false;
            }
        }

        private void btn_Users_Click(object sender, EventArgs e)
        {
            if (operation == "INS")
            {
                if (tb_login.Text != "" && tb_fio.Text != "" && tb_PFM_users.Text != "" && tb_PFMname.Text != "")
                {
                    DataSaver ds = new DataSaver(login_, password_);
                    lb_status.Text = ds.AddUsers
                                                (tb_login.Text.Replace(" ","")
                                                , tb_fio.Text
                                                , tb_PFM_users.Text.Replace(" ", "")
                                                , tb_PFMname.Text
                                                , cb_groupUsers.SelectedValue.ToString()
                                                );
                    
                }
                else
                {
                    lb_status.Text = "Ошибка: Введите всю информацию!";
                }

                tb_login.Text = "";
                tb_fio.Text = "";
                tb_PFM_users.Text = "";
                tb_PFMname.Text = "";
            }
            else if (operation == "UPD")
            {
                if (tb_login.Text != "" && tb_fio.Text != "" && tb_PFM_users.Text != "" && tb_PFMname.Text != "")
                {
                    DataSaver ds = new DataSaver(login_, password_);
                    lb_status.Text = ds.UpdateUsers
                                                (tb_login.Text.Replace(" ", "")
                                                , tb_fio.Text
                                                , tb_PFM_users.Text.Replace(" ", "")
                                                , tb_PFMname.Text
                                                , cb_groupUsers.SelectedValue.ToString()
                                                , tb_users_ID.Text
                                                );
                    LimitsFill();
                    ActualPaymentsFill();
                    ContractsFill();
                    PlansFill();
                }
                else
                {
                    lb_status.Text = "Ошибка: Заполнены не все поля!";
                }
            }
            else if (operation == "DEL")
            {

                DataSaver ds = new DataSaver(login_, password_);
                lb_status.Text = ds.DeleteUsers(tb_users_ID.Text);
            }

            this.v_Users_GETTableAdapter.Fill(this.formationOfScheduleDataSet1.v_Users_GET);
        }

        private void fillByToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.currencyCursTableAdapter1.FillBy(this.currency_.CurrencyCurs);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void button8_Click(object sender, EventArgs e)
        {
            string curs = "";
            if (tb_curs.Text != "")
                curs = tb_curs.Text;
            else
                curs = cb_curs.SelectedValue.ToString();

            if (tb_currency_rub.Text == "")
            {
                lb_status.Text = "Ошибка: Введите стоимость валюты!";
            }

            else
            {
                DataSaver ds = new DataSaver(login_, password_);
                lb_status.Text = ds.AddCurrency(dateTimePicker2.Text
                                                , curs.Replace(" ", "")
                                                , tb_currency_rub.Text.Replace(separator_false, separator_true).Replace(" ","")
                                                );
            }
            this.v_CurrencyCursTableAdapter.Fill(this.currencyCurs1.v_CurrencyCurs);
            this.v_CurrencyTableAdapter.Fill(this.formationOfScheduleDataSet2.v_Currency);

            tb_currency_rub.Text = "";
            tb_curs.Text = "";
        }

        private void CurrencyView_CellValueChanged(object sender, CellValueChangedEventArgs e)
        {
            if (CurrencyView.GetFocusedDataRow()["Rus"].ToString() == "" || CurrencyView.GetFocusedDataRow()["Currency"].ToString() == "")
            {
                lb_status.Text = "Ошибка: Вы оставили пустое поле!";
            }
            else
            {
                DataSaver ds = new DataSaver(login_, password_);
               lb_status.Text = ds.CurrencyCursFill(CurrencyView.GetFocusedDataRow());
            }
        }


        private void Updating()
        {
            this.v_CurrencyCursTableAdapter.Fill(this.currencyCurs1.v_CurrencyCurs);
            this.v_CurrencyTableAdapter.Fill(this.formationOfScheduleDataSet2.v_Currency);
            this.currencyCursTableAdapter1.Fill(this.currency_.CurrencyCurs);
            this.currencyCursTableAdapter.Fill(this.currencyCurs._CurrencyCurs);
            this.usersGroupTableAdapter.Fill(this.fS_UsersGroup.UsersGroup);
            this.v_Users_GETTableAdapter.Fill(this.formationOfScheduleDataSet1.v_Users_GET);
            this.monthTableAdapter.Fill(this.formationOfScheduleDataSet.Month);
            ContractsFill();
            LimitsFill();
            PlansFill();
            ActualPaymentsFill();
            lb_status.Text = "";

            ContractView.FocusedRowHandle = focused_Contracts;
            CurrencyView.FocusedRowHandle = focused_Curs;
            ActualView.FocusedRowHandle = focused_FactPayments;
            limitsView.FocusedRowHandle = focused_limits;
            gridView1.FocusedRowHandle = focused_PaymentsPlan_gridView1;
            gridView2.FocusedRowHandle = focused_PaymentsPlan_gridView2;
            usersView.FocusedRowHandle = focused_users;

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
          //  Updating();
            lb_status.Text = "";
        }

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
           // Updating();
            lb_status.Text = "";
        }

        private void KSSStoContragentFill()
        {
            SqlCommand KSSSComm = new SqlCommand("SELECT [IdKSSStoContragent]" +
                                                        " ,[ContragentType] " +
                                                        " ,[ContragentName]" +
                                                        " ,[KSSScode]" +
                                                        " ,[ContragentCode]" +
                                        " FROM [KSSStoContragent]", _connection);

            try
            {
                if (_connection.State != System.Data.ConnectionState.Open)
                    _connection.Open();
                KSSSComm.ExecuteScalar();

                SqlDataAdapter new_dataAdapter = new SqlDataAdapter(KSSSComm);
                var dataSet = new DataSet();
                new_dataAdapter.Fill(dataSet, "KSSSFill");
                KSSSControl.DataSource = dataSet;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            _connection.Close();
        }

        string operation_ksss = "INS";

        private void link_ksss_add_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link_ksss_add.LinkVisited = true;
            link_ksss_upd.LinkVisited = false;
            link_ksss_Del.LinkVisited = false;
            operation_ksss = "INS";
            btn_ksss.Text = "Добавить";
            gb_ksss.Enabled = true;
        }

        private void link_ksss_upd_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (KSSSView.RowCount != 0)
            {
                link_ksss_add.LinkVisited = false;
                link_ksss_upd.LinkVisited = true;
                link_ksss_Del.LinkVisited = false;
                operation_ksss = "UPD";
                btn_ksss.Text = "Изменить";

                tb_contr_type.Text = KSSSView.GetFocusedDataRow()["ContragentType"].ToString();
                tb_contr_name.Text = KSSSView.GetFocusedDataRow()["ContragentName"].ToString();
                tb_ksss.Text = KSSSView.GetFocusedDataRow()["KSSScode"].ToString();
                tb_contr_code.Text = KSSSView.GetFocusedDataRow()["ContragentCode"].ToString();
                tb_ksss_id.Text = KSSSView.GetFocusedDataRow()["IdKSSStoContragent"].ToString();
                gb_ksss.Enabled = true;
            }
        }
        
        private void link_ksss_Del_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (KSSSView.RowCount != 0)
            {
                link_ksss_add.LinkVisited = false;
                link_ksss_upd.LinkVisited = false;
                link_ksss_Del.LinkVisited = true;
                operation_ksss = "DEL";
                btn_ksss.Text = "Удалить";
                tb_contr_type.Text = KSSSView.GetFocusedDataRow()["ContragentType"].ToString();
                tb_contr_name.Text = KSSSView.GetFocusedDataRow()["ContragentName"].ToString();
                tb_ksss.Text = KSSSView.GetFocusedDataRow()["KSSScode"].ToString();
                tb_contr_code.Text = KSSSView.GetFocusedDataRow()["ContragentCode"].ToString();
                tb_ksss_id.Text = KSSSView.GetFocusedDataRow()["IdKSSStoContragent"].ToString();
                gb_ksss.Enabled = false;
            }
        }

        private void btn_ksss_Click(object sender, EventArgs e)
        {
            int focusRow = 1;

            if (operation_ksss == "UPD")
            {
                if ((tb_contr_type.Text == "" && tb_contr_name.Text == "") ||
                    (tb_ksss.Text == "" && tb_contr_code.Text == ""))
                {
                    lb_status.Text = "Ошибка: Не обязательным для ввода является тольео поле \"Код контрагента/кредитора\"";
                }

                else
                {
                    DataSaver ds = new DataSaver(login_, password_);
                    lb_status.Text = ds.UpdateKSSStoContragent(tb_contr_type.Text, tb_contr_name.Text
                                                , tb_ksss.Text.Replace(" ",""), tb_contr_code.Text.Replace(" ",""), tb_ksss_id.Text);
                }
                focusRow = KSSSView.FocusedRowHandle;

            }

            else if (operation_ksss == "INS")
            {
                if ((tb_contr_type.Text == "" && tb_contr_name.Text == "") ||
                    (tb_ksss.Text == "" && tb_contr_code.Text == ""))
                {
                    lb_status.Text = "Ошибка: Не обязательным для ввода является тольео поле \"Код контрагента/кредитора\"";
                }

                else
                { 
                DataSaver ds = new DataSaver(login_, password_);
                lb_status.Text = ds.InsertKsssContragent(tb_contr_type.Text, tb_contr_name.Text, tb_ksss.Text.Replace(" ",""), tb_contr_code.Text.Replace(" ",""));
                }
                focusRow = KSSSView.RowCount;
                
                tb_contr_code.Text = "";
                tb_contr_name.Text = "";
                tb_contr_type.Text = "";
                tb_ksss.Text = "";
            }

            else if (operation_ksss == "DEL")
            {
                DataSaver ds = new DataSaver(login_, password_);
                lb_status.Text = ds.DelKSSSContragent(tb_ksss_id.Text);
                focusRow = KSSSView.RowCount - 1;
            }
            try
            {
                KSSStoContragentFill();
                KSSSView.FocusedRowHandle = focusRow;
            }
            catch { }

        }

        private void KSSSView_FocusedRowChanged(object sender, FocusedRowChangedEventArgs e)
        {
            if (operation_ksss == "DEL" || operation_ksss == "UPD")
            {
                try
                {
                    tb_contr_type.Text = KSSSView.GetFocusedDataRow()["ContragentType"].ToString();
                    tb_contr_name.Text = KSSSView.GetFocusedDataRow()["ContragentName"].ToString();
                    tb_ksss.Text = KSSSView.GetFocusedDataRow()["KSSScode"].ToString();
                    tb_contr_code.Text = KSSSView.GetFocusedDataRow()["ContragentCode"].ToString();
                    tb_ksss_id.Text = KSSSView.GetFocusedDataRow()["IdKSSStoContragent"].ToString();
                }
                catch { }
            }
        }

        private void btn_limits_load_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpFilD = new OpenFileDialog();
            if (OpFilD.ShowDialog() == DialogResult.OK)
            {
                lb_limits_file.Text = OpFilD.FileName;
                _fileName = OpFilD.FileName;
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += new DoWorkEventHandler(bw_DoWork);
                bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted1);
                groupBox7.Enabled = false;
                Cursor = Cursors.WaitCursor;
                bw.RunWorkerAsync();
            }
        }

        void bw_RunWorkerCompleted1(object sender, RunWorkerCompletedEventArgs e)
        {
            this.lb_limits_sheets.DataSource = null;
            this.lb_limits_sheets.DataSource = _sheets;
            this.groupBox7.Enabled = true;
            Cursor = Cursors.Arrow;
        }

        private void lb_limits_sheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.lb_limits_sheets.DataSource != null)
            {
                _sheet = this.lb_limits_sheets.SelectedValue.ToString();
                char[] charsToTrim = { '$', '#' };
                _sheet = _sheet.Trim(charsToTrim);
            }
        }

        private void btn_limits_get_Click(object sender, EventArgs e)
        {
            if (lb_limits_sheets.Text != "")
            {

                string Warning = "";
                int rows;
                if (int.TryParse(tb_limits_end.Text, out rows) && (Convert.ToInt32(tb_limits_end.Text) - Convert.ToInt32(tb_limits_first.Text)) >= 0)
                {
                    //начать загрузку
                    if (tb_limits_first.Text == "")
                        tb_limits_first.Text = "1";
                    _range = "A" + (Convert.ToInt32(tb_limits_first.Text) - 1).ToString() + ":F" + tb_limits_end.Text;
                    OdbcConnection cn = new OdbcConnection();
                    cn.ConnectionString = string.Format(@"Driver={{Microsoft Excel Driver (*.xls)}};DBQ={0};ReadOnly=0;", lb_limits_file.Text);
                    string strCom = "select * from [" + _sheet + "$" + _range + "]";
                    cn.Open();
                    OdbcCommand comm_mon = new OdbcCommand(strCom, cn);
                    OdbcDataAdapter da = new OdbcDataAdapter();
                    da.SelectCommand = comm_mon;
                    System.Data.DataTable dt = new System.Data.DataTable();
                    da.Fill(dt);
                    DataSaver ds = new DataSaver(login_, password_);
                    Warning = ds.SaveLimits(dt, Convert.ToInt32(tb_limits_first.Text));
                    LimitsFill();

                    if (Warning != "")
                        MessageBox.Show(Warning);
                }
                else
                {
                    MessageBox.Show("Укажите корректное количество строк!");
                }
            }
        }

        private void linkDel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (limitsView.RowCount != 0)
            {
                link = "DEL";
                linkAdd.LinkVisited = false;
                linkChenge.LinkVisited = false;
                linkDel.LinkVisited = true;
                btnAdd.Text = "Удалить";
                tb_PFM.Text = limitsView.GetFocusedDataRow()["lim1"].ToString();
                tb_finpos.Text = limitsView.GetFocusedDataRow()["lim2"].ToString();
                tb_year.Text = limitsView.GetFocusedDataRow()["lim3"].ToString();
                tb_summ.Text = limitsView.GetFocusedDataRow()["lim5"].ToString();
                cb_month.SelectedValue = limitsView.GetFocusedDataRow()["limMonth"].ToString();
                tb_id.Text = limitsView.GetFocusedDataRow()["limID"].ToString();
                groupBox3.Enabled = false;
            }
        }

        private void btn_actual_load_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpFilD = new OpenFileDialog();
            if (OpFilD.ShowDialog() == DialogResult.OK)
            {
                lb_actual_file.Text = OpFilD.FileName;
                _fileName = OpFilD.FileName;
                BackgroundWorker bw = new BackgroundWorker();
                bw.DoWork += new DoWorkEventHandler(bw_DoWork);
                bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted2);
                gb_actual.Enabled = false;
                Cursor = Cursors.WaitCursor;
                bw.RunWorkerAsync();
            }
        }

        void bw_RunWorkerCompleted2(object sender, RunWorkerCompletedEventArgs e)
        {
            this.lb_actual_sheets.DataSource = null;
            this.lb_actual_sheets.DataSource = _sheets;
            this.gb_actual.Enabled = true;
            Cursor = Cursors.Arrow;
        }

        private void btn_actual_insert_Click(object sender, EventArgs e)
        {
            if (lb_actual_sheets.Text != "")
            {
                string Warning = "";
                int rows;
                if (int.TryParse(tb_actual_end.Text, out rows))
                {
                    //начать загрузку
                    if (tb_actual_first.Text == "")
                        tb_actual_first.Text = "1";
                    _range = "A" + (Convert.ToInt32(tb_actual_first.Text) - 1).ToString() + ":H" + tb_actual_end.Text;
                    OdbcConnection cn = new OdbcConnection();
                    cn.ConnectionString = string.Format(@"Driver={{Microsoft Excel Driver (*.xls)}};DBQ={0};ReadOnly=0;", lb_actual_file.Text);
                    string strCom = "select * from [" + _sheet + "$" + _range + "]";
                    cn.Open();
                    OdbcCommand comm_mon = new OdbcCommand(strCom, cn);
                    OdbcDataAdapter da = new OdbcDataAdapter();
                    da.SelectCommand = comm_mon;
                    System.Data.DataTable dt = new System.Data.DataTable();
                    da.Fill(dt);
                    DataSaver ds = new DataSaver(login_, password_);
                    Warning = ds.SaveActualPayment(dt, Convert.ToInt32(tb_actual_first.Text));
                    ActualPaymentsFill();
                    if (Warning != "")
                        MessageBox.Show(Warning);

                }
                else
                {
                    MessageBox.Show("Укажите количество строк");
                }
            }
        }

        private void lb_actual_sheets_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (this.lb_actual_sheets.DataSource != null)
            {
                _sheet = this.lb_actual_sheets.SelectedValue.ToString();
                char[] charsToTrim = { '$', '#' };
                _sheet = _sheet.Trim(charsToTrim);
            }
        }

        private void link_Payment_add_W_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (operation_pay == "INS")
            {
                lb_status.Text = "";
            tb_pay_ContragentCode.Text = gridView1.GetFocusedDataRow()["col1"].ToString();
            tp_pay_ContractCode.Text = gridView1.GetFocusedDataRow()["col2"].ToString();
            cb_pay_FinPosition.SelectedValue = gridView1.GetFocusedDataRow()["col4"].ToString();
            tb_pay_Summ.Text = "";
            dtp_pay_Date.Value = DateTime.Today;
            link_Payment_add_W.LinkVisited = true;
            FinPositionChenged();
            }
        }

        string operation_pay = "INS";
        private void link_payment_add_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lb_status.Text = "";
            operation_pay = "INS";
            link_payment_add.LinkVisited = true;
            link_payment_update.LinkVisited = false;
            link_payment_del.LinkVisited = false;
            gb_pay1.Enabled = true;
            gb_pay2.Enabled = true;
            button7.Text = "Добавить";
            tb_pay_PFMcode.Text = inform_[1];
            link_Payment_add_W.Enabled = true;
            tp_pay_ContractCode.Text = "";
            tb_pay_ContragentCode.Text = "";
            dtp_pay_Date.Value = DateTime.Today;
            tb_pay_Summ.Text = "";
            tb_Comment.Text = "";
            FinPositionChenged();                                          
            
        }

        private void link_payment_update_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (gridView2.RowCount != 0)
            {
                lb_status.Text = "";
                operation_pay = "UPD";
                link_payment_add.LinkVisited = false;
                link_payment_update.LinkVisited = true;
                link_payment_del.LinkVisited = false;
                button7.Text = "Изменить";
                tb_pay_ContragentCode.Text = gridView2.GetFocusedDataRow()["plan2"].ToString();
                tp_pay_ContractCode.Text = gridView2.GetFocusedDataRow()["plan1"].ToString();
                // tb_pay_EPL.Text = gridView2.GetFocusedDataRow()["plan5"].ToString();
                tb_pay_PFMcode.Text = gridView2.GetFocusedDataRow()["plan6"].ToString();
                tb_pay_ContragentType.Text = gridView2.GetFocusedDataRow()["plan3"].ToString();
                cb_pay_FinPosition.SelectedValue = gridView2.GetFocusedDataRow()["plan4"].ToString();
                dtp_pay_Date.Text = gridView2.GetFocusedDataRow()["plan7"].ToString();
                cb_pay_Currency.SelectedValue = gridView2.GetFocusedDataRow()["plan8"].ToString();
                tb_pay_Summ.Text = gridView2.GetFocusedDataRow()["plan9"].ToString();
                tb_pay_SummRUB.Text = gridView2.GetFocusedDataRow()["plan10"].ToString();
                tb_pay_id.Text = gridView2.GetFocusedDataRow()["planID"].ToString();
                tb_plan_statelikvid.Text = gridView2.GetFocusedDataRow()["plan12"].ToString();
                tb_plan_stavrolen.Text = gridView2.GetFocusedDataRow()["plan11"].ToString();
                tb_pay_finpos_copy.Text = gridView2.GetFocusedDataRow()["plan4"].ToString();
                tb_Comment.Text = gridView2.GetFocusedDataRow()["planComment"].ToString();
                link_Payment_add_W.Enabled = false;
                gb_pay1.Enabled = true;
                gb_pay2.Enabled = true;
                FinPositionChenged();
            }
        }

        private void link_payment_del_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (gridView2.RowCount != 0)
            {
                try
                {
                    lb_status.Text = "";
                    operation_pay = "DEL";
                    link_payment_add.LinkVisited = false;
                    link_payment_update.LinkVisited = false;
                    link_payment_del.LinkVisited = true;
                    button7.Text = "Удалить";
                    tb_pay_ContragentCode.Text = gridView2.GetFocusedDataRow()["plan2"].ToString();
                    tp_pay_ContractCode.Text = gridView2.GetFocusedDataRow()["plan1"].ToString();
                    // tb_pay_EPL.Text = gridView2.GetFocusedDataRow()["plan5"].ToString();
                    tb_pay_PFMcode.Text = gridView2.GetFocusedDataRow()["plan6"].ToString();
                    tb_pay_ContragentType.Text = gridView2.GetFocusedDataRow()["plan3"].ToString();
                    cb_pay_FinPosition.SelectedValue = gridView2.GetFocusedDataRow()["plan4"].ToString();
                    dtp_pay_Date.Text = gridView2.GetFocusedDataRow()["plan7"].ToString();
                    cb_pay_Currency.SelectedValue = gridView2.GetFocusedDataRow()["plan8"].ToString();
                    tb_pay_Summ.Text = gridView2.GetFocusedDataRow()["plan9"].ToString();
                    tb_pay_SummRUB.Text = gridView2.GetFocusedDataRow()["plan10"].ToString();
                    tb_pay_id.Text = gridView2.GetFocusedDataRow()["planID"].ToString();
                    tb_plan_statelikvid.Text = gridView2.GetFocusedDataRow()["plan12"].ToString();
                    tb_plan_stavrolen.Text = gridView2.GetFocusedDataRow()["plan11"].ToString();
                    tb_pay_finpos_copy.Text = gridView2.GetFocusedDataRow()["plan4"].ToString();
                    tb_Comment.Text = gridView2.GetFocusedDataRow()["planComment"].ToString();
                    link_Payment_add_W.Enabled = false;
                    gb_pay1.Enabled = false;
                    gb_pay2.Enabled = false;
                    FinPositionChenged();
                }
                catch { }
            }
        }
        int focused_PaymentsPlan_gridView2 = 0;
        private void gridView2_FocusedRowChanged(object sender, FocusedRowChangedEventArgs e)
        {
            
            focused_PaymentsPlan_gridView2 = gridView2.FocusedRowHandle;
            if (operation_pay == "UPD" || operation_pay == "DEL")
            {
                try
                {
                    tb_pay_ContragentCode.Text = gridView2.GetFocusedDataRow()["plan2"].ToString();
                    tp_pay_ContractCode.Text = gridView2.GetFocusedDataRow()["plan1"].ToString();
                    //tb_pay_EPL.Text = gridView2.GetFocusedDataRow()["plan5"].ToString();
                    tb_pay_PFMcode.Text = gridView2.GetFocusedDataRow()["plan6"].ToString();
                    tb_pay_ContragentType.Text = gridView2.GetFocusedDataRow()["plan3"].ToString();
                    cb_pay_FinPosition.SelectedValue = gridView2.GetFocusedDataRow()["plan4"].ToString();

                    dtp_pay_Date.Text = gridView2.GetFocusedDataRow()["plan7"].ToString();
                    cb_pay_Currency.SelectedValue = gridView2.GetFocusedDataRow()["plan8"].ToString();
                    tb_pay_Summ.Text = gridView2.GetFocusedDataRow()["plan9"].ToString();
                    tb_pay_SummRUB.Text = gridView2.GetFocusedDataRow()["plan10"].ToString();
                    tb_pay_id.Text = gridView2.GetFocusedDataRow()["planID"].ToString();
                    tb_plan_statelikvid.Text = gridView2.GetFocusedDataRow()["plan12"].ToString();
                    tb_plan_stavrolen.Text = gridView2.GetFocusedDataRow()["plan11"].ToString();
                    tb_pay_finpos_copy.Text = gridView2.GetFocusedDataRow()["plan4"].ToString();
                    tb_Comment.Text = gridView2.GetFocusedDataRow()["planComment"].ToString();
                    FinPositionChenged();
                }
                catch { }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string answer = "";
            if (dtp_pay_Date.Value == DateTime.Today)
            {
                string mess = "Вы уверенны, что платеж должен быть совершен сегодня?";
                string caption = "Вопрос";
                var result = MessageBox.Show(mess, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    answer = "Yes";
                }
                else answer = "No";
            }
            if (answer != "No")
            {

                //   if (dtp_pay_Date.Value == DateTime.Today)
                //      MessageBox.Show("Вы уверенны, что платеж должен быть совершен текущим днем?");

                int focused_row = 0;
                focused_row = gridView2.FocusedRowHandle;
                DataSaver ds1 = new DataSaver(login_, password_);
                inform_ = ds1.getPFM();

                string admMet = "";
                //if (inform_[0].Substring(0, 13) == "Администратор") Изменил Дороненков Г.Г. 03-06-2014 
				if ((inform_[0] != string.Empty) && (inform_[0].Substring(0, 13) == "Администратор"))
                    admMet = "0";
                else
                    admMet = "1";

                if (inform_[2] == "1" && admMet == "1")
                {
                    button7.Enabled = false;
                    lBlockMessage.Visible = true;

                    tb_pay_ContragentCode.Text = "";
                    tb_pay_Summ.Text = "";
                    tp_pay_ContractCode.Text = "";
                    dtp_pay_Date.Value = DateTime.Today;
                    tb_plan_statelikvid.Text = "";
                    tb_plan_stavrolen.Text = "";
                    tb_pay_EPL.Text = "";
                    tb_Comment.Text = "";
                }
                else
                {

                    string curs = "";
                    curs = cb_pay_Currency.SelectedValue.ToString();
                    try
                    {
                        if (operation_pay == "INS")
                        {
                            string message = "";
                            DataSaver ds = new DataSaver(login_, password_);
                            message = ds.InsertPaymentPlan(tp_pay_ContractCode.Text,
                                                                  tb_pay_ContragentCode.Text,
                                                                  tb_pay_ContragentType.Text,
                                                                  cb_pay_FinPosition.SelectedValue.ToString(),
                                                                  tb_pay_EPL.Text,
                                                                  tb_pay_PFMcode.Text.Replace(" ", ""),
                                                                  dtp_pay_Date.Text,
                                                                  curs,
                                                                  tb_pay_Summ.Text.Replace(" ", ""),
                                                                  tb_pay_SummRUB.Text,
                                                                  tb_Comment.Text
                                                                  );

                            if (message.Substring(0, 14) == "Предупреждение" || message.Substring(0, 5) == "Лимит")
                                MessageBox.Show(message);
                            else lb_status.Text = message;

                            tb_pay_ContragentCode.Text = "";
                            tb_pay_Summ.Text = "";
                            tp_pay_ContractCode.Text = "";
                            dtp_pay_Date.Value = DateTime.Today;
                            tb_plan_statelikvid.Text = "";
                            tb_plan_stavrolen.Text = "";
                            tb_pay_EPL.Text = "";
                            tb_Comment.Text = "";
                            focused_row = gridView2.RowCount;
                        }

                        else if (operation_pay == "UPD")
                        {
                            focused_row = gridView2.FocusedRowHandle;
                            DataSaver ds = new DataSaver(login_, password_);
                            lb_status.Text = ds.UpdatePaymentPlan(tp_pay_ContractCode.Text,
                                                                  tb_pay_ContragentCode.Text,
                                                                  tb_pay_ContragentType.Text,
                                                                  cb_pay_FinPosition.SelectedValue.ToString(),
                                                                  tb_pay_EPL.Text,
                                                                  tb_pay_PFMcode.Text.Replace(" ", ""),
                                                                  dtp_pay_Date.Text,
                                                                  curs,
                                                                  tb_pay_Summ.Text.Replace(" ", ""),
                                                                  tb_pay_SummRUB.Text,
                                                                  tb_pay_id.Text,
                                                                  tb_Comment.Text,
                                                                  admMet
                                                                  );
                        }

                        else if (operation_pay == "DEL")
                        {
                            try
                            {
                                if (gridView2.FocusedRowHandle != gridView2.RowCount - 1)
                                    focused_row = gridView2.FocusedRowHandle;
                                else
                                    focused_row = gridView2.FocusedRowHandle - 1;
                            }
                            catch { }
                            DataSaver ds = new DataSaver(login_, password_);
                            lb_status.Text = ds.DeletePayments(tb_pay_id.Text);

                        }
                    }
                    catch
                    {
                        lb_status.Text = "Ошибка: Заполнены не все поля!!!";
                    }
                    PlansFill();
                    gridView2.FocusedRowHandle = focused_row;
                    link_Payment_add_W.LinkVisited = false;
                }
            }
        }

        private void cb_pay_FinPosition_SelectedIndexChanged(object sender, EventArgs e)
        {
            FinPositionChenged();
        }

        public void FinPositionChenged()
        {
            try
            {
                tb_pay_finpos_copy.Text = cb_pay_FinPosition.SelectedValue.ToString();
                DataSaver ds = new DataSaver(login_, password_);
                tb_pay_EPL.Text = ds.GetEPL(tb_pay_finpos_copy.Text);

                DataSaver ds1 = new DataSaver(login_, password_);
                tb_plan_stavrolen.Text = ds1.GetFinPositonName(tb_pay_finpos_copy.Text);
            }
            catch { }
        }

        private void tb_pay_EPL_TextChanged(object sender, EventArgs e)
        {
            try
            {
                DataSaver ds = new DataSaver(login_, password_);
                tb_plan_statelikvid.Text = ds.GetStateEverydayLicvid(tb_pay_EPL.Text);
            }
            catch { }
        }


        private void tb_pay_ContragentCode_TextChanged(object sender, EventArgs e)
        {
            DataSaver ds = new DataSaver(login_, password_);
            tb_pay_ContragentType.Text = ds.GetContragentType(tb_pay_ContragentCode.Text);

            if (tb_pay_ContragentType.Text.Substring(0,26) != "Контрагента не существует!")
            {
                tp_pay_ContractCode.Enabled = true;
                cb_pay_FinPosition.Enabled = true;
                gb_pay2.Enabled = true;
                link_payment_add.Enabled = true;
                link_payment_del.Enabled = true;
                link_payment_update.Enabled = true;
                tb_pay_EPL.Enabled = true;
                tb_pay_PFMcode.Enabled = true;
                tb_pay_SummRUB.Enabled = true;
                if(inform_[2] != "1")
                    button7.Enabled = true;
                tb_plan_statelikvid.Enabled = true;
                tb_plan_stavrolen.Enabled = true;
                tb_pay_ContragentType.ForeColor = System.Drawing.Color.Black;
            }

            else
            {
                tp_pay_ContractCode.Enabled = false;
                cb_pay_FinPosition.Enabled = false;
                gb_pay2.Enabled = false;
                link_payment_add.Enabled = false;
                link_payment_del.Enabled = false;
                link_payment_update.Enabled = false;
                button7.Enabled = false;
                tb_pay_EPL.Enabled = false;
                tb_pay_PFMcode.Enabled = false;
                tb_pay_SummRUB.Enabled = false;
                tb_plan_statelikvid.Enabled = false;
                tb_plan_stavrolen.Enabled = false;
                tb_pay_ContragentType.ForeColor = System.Drawing.Color.Red;

            }

        }

        private void tb_pay_Summ_TextChanged(object sender, EventArgs e)
        {
            try
            {
                DataSaver ds = new DataSaver(login_, password_);
                tb_pay_SummRUB.Text = ds.GetRUSsumm(tb_pay_Summ.Text
                                                    , cb_pay_Currency.SelectedValue.ToString()
                                                    , dtp_pay_Date.Value.Date.ToString()
                                                    );
            }
            catch { }
        }

        private void gridView3_CustomDrawCell(object sender, RowCellCustomDrawEventArgs e)
        {
            if (e.RowHandle == gridView3.FocusedRowHandle)
            {
                e.Appearance.BackColor = Color.FromArgb(0, 128, 255);
                e.Appearance.ForeColor = Color.White;
            }
            else
            {

                if (gridView3.GetDataRow(e.RowHandle)["FPcodeStavrolen"].ToString() == "")
                    e.Appearance.BackColor = Color.FromArgb(255, 255, 220);
            }
        }

        string operation_fp = "INS";

        private void link_fp_delete_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (gridView11.RowCount != 0)
            {
                btn_fp_EPL.Text = "Удалить";
                btn_fp_Stavrolen.Text = "Удалить";
                link_fp_delete.LinkVisited = true;
                link_fp_insert.LinkVisited = false;
                link_fp_update.LinkVisited = false;
                operation_fp = "DEL";
            }
        }

        private void link_fp_insert_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            btn_fp_EPL.Text = "Добавить";
            btn_fp_Stavrolen.Text = "Добавить";
            link_fp_delete.LinkVisited = false;
            link_fp_insert.LinkVisited = true;
            link_fp_update.LinkVisited = false;
            operation_fp = "INS";
        }

        private void linkfp_update_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (gridView11.RowCount != 0)
            {
                btn_fp_EPL.Text = "Изменить";
                btn_fp_Stavrolen.Text = "Изменить";
                link_fp_delete.LinkVisited = false;
                link_fp_insert.LinkVisited = false;
                link_fp_update.LinkVisited = true;
                operation_fp = "UPD";
                tb_fp_FPcode.Text = gridView11.GetFocusedDataRow()["FPcode"].ToString();
                tb_fp_Licvid.Text = gridView11.GetFocusedDataRow()["StateEverydayLicvid"].ToString();
                FPCODESTAVROLEN =
                tb_fp_FPstavrolen.Text = gridView12.GetFocusedDataRow()["FPcodeStavrolen"].ToString();
                tb_fp_FPSname.Text = gridView12.GetFocusedDataRow()["FPStavrolenName"].ToString();
            }
        }

        string FPCODE_save = "";

        private void gridView11_FocusedRowChanged(object sender, FocusedRowChangedEventArgs e)
        {
            try
            {
                string FP_EPL = "";

                FP_EPL = gridView11.GetFocusedDataRow()["FPcode"].ToString();

                string comm = "Select IdFinPosition, FPcodeStavrolen, FPStavrolenName FROM udf_FS_Get_FPStavrolen(" + FP_EPL + ") ";
                SqlCommand Comm = new SqlCommand(comm, _connection);
                if (_connection.State != System.Data.ConnectionState.Open)
                    _connection.Open();
                Comm.ExecuteScalar();

                SqlDataAdapter new_dataAdapter = new SqlDataAdapter(Comm);
                var dataSet = new DataSet();
                new_dataAdapter.Fill(dataSet, "FPFill");
                gridControl6.DataSource = dataSet;
                _connection.Close();
                
                if (operation_fp == "UPD" || operation_fp == "DEL")
            {
                    FPCODE_save = 
                tb_fp_FPcode.Text = gridView11.GetFocusedDataRow()["FPcode"].ToString();
                tb_fp_Licvid.Text = gridView11.GetFocusedDataRow()["StateEverydayLicvid"].ToString();
            }

            }
            catch
            { 
            
            }


        }

        private void btn_fp_EPL_Click(object sender, EventArgs e)
        {
            if (operation_fp == "INS")
            {
                DataSaver ds = new DataSaver(login_, password_);
                lb_status.Text = ds.InsertFinPosition(tb_fp_FPcode.Text, tb_fp_Licvid.Text);

                tb_fp_FPcode.Text = "";
                tb_fp_Licvid.Text = "";
                
            }

            else if (operation_fp == "UPD")
            {
                try
                {
                   // string ID = gridView12.GetFocusedDataRow()["IdFinPosition"].ToString();
                    DataSaver ds = new DataSaver(login_, password_);
                    lb_status.Text = ds.UpdateFinPosition(FPCODE_save,tb_fp_FPcode.Text, tb_fp_Licvid.Text);
                }
                catch { }
            }

            else if (operation_fp == "DEL")
            {
                DataSaver ds = new DataSaver(login_, password_);
                lb_status.Text = ds.DeleteFinPosition(tb_fp_FPcode.Text, tb_fp_Licvid.Text);
            }
            this.v_FinPosition_EPLTableAdapter.Fill(this.formationOfScheduleDataSet5.v_FinPosition_EPL);
            this.financialPositionTableAdapter.Fill(this.formationOfScheduleDataSet3.FinancialPosition);
            this.v_FinancialPosition_TopicTableAdapter.Fill(this.formationOfScheduleDataSet6.v_FinancialPosition_Topic);
            ContractsFill();
            ActualPaymentsFill();
            PlansFill();
        }

        private void btn_fp_Stavrolen_Click(object sender, EventArgs e)
        {
            if (operation_fp == "INS")
            {
                string FPcode = "";
                FPcode = gridView11.GetFocusedDataRow()["FPcode"].ToString();
                DataSaver ds = new DataSaver(login_, password_);
                lb_status.Text = ds.InsertFinPosition_Stavrolen(FPcode, tb_fp_FPstavrolen.Text, tb_fp_FPSname.Text);

                this.v_FinPosition_EPLTableAdapter.Fill(this.formationOfScheduleDataSet5.v_FinPosition_EPL);

                tb_fp_FPstavrolen.Text = "";
                tb_fp_FPSname.Text = "";
            }

            else if (operation_fp == "UPD")
            { 
                try
                {
                    DataSaver ds = new DataSaver(login_, password_);
                    lb_status.Text = ds.UpdateFinPosition_Stavrolen(FPCODESTAVROLEN
                                                    , tb_fp_FPstavrolen.Text, tb_fp_FPSname.Text);
                }
                catch { }
            }

            else if (operation_fp == "DEL")
            {

                DataSaver ds = new DataSaver(login_, password_);
                lb_status.Text = ds.DeleteFinPosition_Stavrolen(tb_fp_FPstavrolen.Text, tb_fp_FPSname.Text);
            }
           // this.v_FinPosition_EPLTableAdapter.Fill(this.formationOfScheduleDataSet5.v_FinPosition_EPL);
            this.financialPositionTableAdapter.Fill(this.formationOfScheduleDataSet3.FinancialPosition);
            this.v_FinancialPosition_TopicTableAdapter.Fill(this.formationOfScheduleDataSet6.v_FinancialPosition_Topic);
            ContractsFill();
            ActualPaymentsFill();
            PlansFill();
        }

        private void btn_UP_Click(object sender, EventArgs e)
        {
            int focusRow = gridView3.FocusedRowHandle;
            string id_tec = gridView3.GetFocusedDataRow()["IdFinPosition"].ToString();
            gridView3.MovePrev();
            string id_up = gridView3.GetFocusedDataRow()["IdFinPosition"].ToString();

            DataSaver ds = new DataSaver(login_, password_);
            lb_status.Text = ds.SortFinPos(id_tec,id_up);

            this.v_FinancialPosition_TopicTableAdapter.Fill(this.formationOfScheduleDataSet6.v_FinancialPosition_Topic);
            gridView3.FocusedRowHandle = focusRow;
        }

        private void btn_DOWN_Click(object sender, EventArgs e)
        {
            int focusRow = gridView3.FocusedRowHandle;
            string id_tec = gridView3.GetFocusedDataRow()["IdFinPosition"].ToString();
            gridView3.MoveNext();
            string id_down = gridView3.GetFocusedDataRow()["IdFinPosition"].ToString();

            DataSaver ds = new DataSaver(login_, password_);
            lb_status.Text = ds.SortFinPos(id_tec, id_down);

            this.v_FinancialPosition_TopicTableAdapter.Fill(this.formationOfScheduleDataSet6.v_FinancialPosition_Topic);
            gridView3.FocusedRowHandle = focusRow;
        }

        private void cb_pay_Currency_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                DataSaver ds = new DataSaver(login_, password_);
                tb_pay_SummRUB.Text = ds.GetRUSsumm(tb_pay_Summ.Text
                                                    , cb_pay_Currency.SelectedValue.ToString()
                                                    ,dtp_pay_Date.Value.Date.ToString()
                                                    );
            }
            catch { }
        }
        string FPCODESTAVROLEN = "";
        private void gridView12_FocusedRowChanged(object sender, FocusedRowChangedEventArgs e)
        {
            try
            {

                if (operation_fp == "UPD" || operation_fp == "DEL")
                {
                    FPCODESTAVROLEN =
                    tb_fp_FPstavrolen.Text = gridView12.GetFocusedDataRow()["FPcodeStavrolen"].ToString();
                    tb_fp_FPSname.Text = gridView12.GetFocusedDataRow()["FPStavrolenName"].ToString();
                }

            }
            catch
            {

            }
        }

        private void gridView3_CustomColumnSort(object sender, CustomColumnSortEventArgs e)
        {
            e.Handled = true;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            _connection.Close();
            temp.Close();
            timer1.Stop();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ReportDeclarationForm declar_frm = new ReportDeclarationForm(_connection, inform_, TemplatePath);
            declar_frm.Show();
        }

            private string CellName(int CellRow, int CellColumn)
            {
                return ((char)(64 + CellColumn) + CellRow.ToString());
            }

            private void Otch3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
            {
                ReportLimits rep_frm = new ReportLimits(_connection, 1, TemplatePath);
                rep_frm.Show();
            }

            private void otch2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
            {
                ReportDate rep_frm = new ReportDate(_connection, 1,  TemplatePath);
                rep_frm.Show();
            }

            private void otch_sv_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
            {
                ReportDate rep_frm = new ReportDate(_connection, 2, TemplatePath);
                rep_frm.Show();
            }

            private void otch5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
            {
                ReportLimits rep_frm = new ReportLimits(_connection, 2, TemplatePath);
                rep_frm.Show();
            }

            private void btn_del_contracts_Click(object sender, EventArgs e)
            {
                int focusRow =  ContractView.FocusedRowHandle;
                    string idRow = "";
                    idRow = ContractView.GetFocusedDataRow()["ID"].ToString();
                    DataSaver ds = new DataSaver(login_, password_);
                    lb_status.Text = ds.DeleteContracts(idRow);
                ContractsFill();
                try
                {
                    if (focusRow != ContractView.RowCount - 1)
                        ContractView.FocusedRowHandle = focusRow - 1;
                    else
                        ContractView.FocusedRowHandle = focusRow;
                }
                catch { }
                _connection.Close();
            }
            int focused_PaymentsPlan_gridView1 = 1;
            private void gridView1_FocusedRowChanged(object sender, FocusedRowChangedEventArgs e)
            {
                lb_status.Text = "";
                focused_PaymentsPlan_gridView1 = gridView1.FocusedRowHandle;
            }
            int focused_FactPayments = 1;
            private void ActualView_FocusedRowChanged(object sender, FocusedRowChangedEventArgs e)
            {
                focused_FactPayments = ActualView.FocusedRowHandle;
            }

            int focused_Contracts = 1;
            private void ContractView_FocusedRowChanged(object sender, FocusedRowChangedEventArgs e)
            {
                focused_Contracts = ContractView.FocusedRowHandle;
            }

            int focused_Curs = 1;
            private void CurrencyView_FocusedRowChanged(object sender, FocusedRowChangedEventArgs e)
            {
                focused_Curs = CurrencyView.FocusedRowHandle;
            }

            private void ErrorList(string error)
            {
                string TemplateFileName = "ErrorList.txt";
                string TemplatePath = Directory.GetCurrentDirectory() + "\\Reports\\" + TemplateFileName;
                StreamWriter sw = new StreamWriter(TemplatePath);
                sw.WriteLine(error);
                sw.Close(); 
                Process.Start(TemplatePath); 

            }

            private void button3_Click(object sender, EventArgs e)
            {
                _range = "A" + (Convert.ToInt32(textBox1.Text) - 1).ToString() + ":J" + textBox2.Text;
                OdbcConnection cn = new OdbcConnection();
                cn.ConnectionString = string.Format(@"Driver={{Microsoft Excel Driver (*.xls)}};DBQ={0};ReadOnly=0;", label1.Text);
                string strCom = "select * from [" + _sheet + "$" + _range + "]";
                cn.Open();
                OdbcCommand comm_mon = new OdbcCommand(strCom, cn);
                OdbcDataAdapter da = new OdbcDataAdapter();
                da.SelectCommand = comm_mon;
                System.Data.DataTable dt = new System.Data.DataTable();
                da.Fill(dt);
                DataSaver ds = new DataSaver(login_, password_);

                ds.SavePaymentsPlan(dt);

            }

            private void button4_Click_1(object sender, EventArgs e)
            {
               
                // TODO: This line of code loads data into the 'formationOfScheduleDataSet6.v_FinancialPosition_Topic' table. You can move, or remove it, as needed.
                this.v_FinancialPosition_TopicTableAdapter.Fill(this.formationOfScheduleDataSet6.v_FinancialPosition_Topic);
                // TODO: This line of code loads data into the 'formationOfScheduleDataSet5.v_FinPosition_EPL' table. You can move, or remove it, as needed.
                this.v_FinPosition_EPLTableAdapter.Fill(this.formationOfScheduleDataSet5.v_FinPosition_EPL);
                // TODO: This line of code loads data into the 'udf_getcurrency.udf_FS_get_currency' table. You can move, or remove it, as needed.
                this.udf_FS_get_currencyTableAdapter.Fill(this.udf_getcurrency.udf_FS_get_currency);
                // TODO: This line of code loads data into the 'formationOfScheduleDataSet4.v_FPcodeStavrolen' table. You can move, or remove it, as needed.
                this.v_FPcodeStavrolenTableAdapter.Fill(this.formationOfScheduleDataSet4.v_FPcodeStavrolen);
                // TODO: This line of code loads data into the 'formationOfScheduleDataSet3.FinancialPosition' table. You can move, or remove it, as needed.
                this.financialPositionTableAdapter.Fill(this.formationOfScheduleDataSet3.FinancialPosition);
                // TODO: This line of code loads data into the 'currencyCurs1.v_CurrencyCurs' table. You can move, or remove it, as needed.
                this.v_CurrencyCursTableAdapter.Fill(this.currencyCurs1.v_CurrencyCurs);
                // TODO: This line of code loads data into the 'formationOfScheduleDataSet2.v_Currency' table. You can move, or remove it, as needed.
                this.v_CurrencyTableAdapter.Fill(this.formationOfScheduleDataSet2.v_Currency);
                // TODO: This line of code loads data into the 'currency_.CurrencyCurs' table. You can move, or remove it, as needed.
                this.currencyCursTableAdapter1.Fill(this.currency_.CurrencyCurs);
                // TODO: This line of code loads data into the 'currencyCurs._CurrencyCurs' table. You can move, or remove it, as needed.
                this.currencyCursTableAdapter.Fill(this.currencyCurs._CurrencyCurs);
                // TODO: This line of code loads data into the 'fS_UsersGroup.UsersGroup' table. You can move, or remove it, as needed.
                this.usersGroupTableAdapter.Fill(this.fS_UsersGroup.UsersGroup);
                // TODO: This line of code loads data into the 'formationOfScheduleDataSet1.v_Users_GET' table. You can move, or remove it, as needed.
                this.v_Users_GETTableAdapter.Fill(this.formationOfScheduleDataSet1.v_Users_GET);
                // TODO: This line of code loads data into the 'formationOfScheduleDataSet.Month' table. You can move, or remove it, as needed.
                this.monthTableAdapter.Fill(this.formationOfScheduleDataSet.Month);
                ContractsFill();
                LimitsFill();
                PlansFill();
                ActualPaymentsFill();
                KSSStoContragentFill();
                CurrencyView.Columns[0].SortOrder = DevExpress.Data.ColumnSortOrder.Descending;
                lb_status.Text = "";
            }

            private void otch6_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
            {
                ReportMark rep_frm = new ReportMark(_connection, TemplatePath, "ReportMark");
                rep_frm.Show();
            }

            private void tabPage3_Click(object sender, EventArgs e)
            {

            }
            string blockMark;
            private void btn_block_Click(object sender, EventArgs e)
            {
                lb_status.Text = "";
                if (blockMark == "lock")
                {
                    string comm = "Update BlockMark SET block = 1";
                    SqlCommand Comm = new SqlCommand(comm, _connection);
                    if (_connection.State != System.Data.ConnectionState.Open)
                        _connection.Open();
                    Comm.ExecuteScalar();
                    btn_block.Text = "Разблокировать ввод ПП";
                    blockMark = "unlock";
                }
                else
                {
                    string comm = "Update BlockMark SET block = 0";
                    SqlCommand Comm = new SqlCommand(comm, _connection);
                    if (_connection.State != System.Data.ConnectionState.Open)
                        _connection.Open();
                    Comm.ExecuteScalar();
                    btn_block.Text = "Заблокировать ввод ПП";
                    blockMark = "lock";
                }

            }

            private void dtp_pay_Date_ValueChanged(object sender, EventArgs e)
            {
                try
                {
                    DataSaver ds = new DataSaver(login_, password_);
                    tb_pay_SummRUB.Text = ds.GetRUSsumm(tb_pay_Summ.Text
                                                        , cb_pay_Currency.SelectedValue.ToString()
                                                        , dtp_pay_Date.Value.Date.ToString()
                                                        );
                }
                catch { }
            }

            private void btn_KSSS_load_Main_Click(object sender, EventArgs e)
            {
                OpenFileDialog OpFilD = new OpenFileDialog();
                if (OpFilD.ShowDialog() == DialogResult.OK)
                {
                    lbl_KSSS_dateload.Text = OpFilD.FileName;
                    _fileName = OpFilD.FileName;
                    BackgroundWorker bw = new BackgroundWorker();
                    bw.DoWork += new DoWorkEventHandler(bw_DoWork);
                    bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted3);
                    groupBox2.Enabled = false;
                    Cursor = Cursors.WaitCursor;
                    bw.RunWorkerAsync();
                }
            }


            void bw_RunWorkerCompleted3(object sender, RunWorkerCompletedEventArgs e)
            {
                this.listBox_KSSS.DataSource = null;
                this.listBox_KSSS.DataSource = _sheets;
                this.groupBox2.Enabled = true;
                Cursor = Cursors.Arrow;
            }

            private void btn_KSSS_load_Click(object sender, EventArgs e)
            {
                if (listBox_KSSS.Text != "")
                {

                    string Warning = "";
                    int rows;
                    if (int.TryParse(tb_KSSS_lastRow.Text, out rows) && (Convert.ToInt32(tb_KSSS_lastRow.Text) - Convert.ToInt32(tb_KSSS_firstRow.Text)) >= 0)
                    {
                        //начать загрузку
                        if (tb_KSSS_firstRow.Text == "")
                            tb_KSSS_firstRow.Text = "1";
                        _range = "A" + (Convert.ToInt32(tb_KSSS_firstRow.Text) - 1).ToString() + ":D" + tb_KSSS_lastRow.Text;
                        OdbcConnection cn = new OdbcConnection();
                        cn.ConnectionString = string.Format(@"Driver={{Microsoft Excel Driver (*.xls)}};DBQ={0};ReadOnly=0;", lbl_KSSS_dateload.Text);
                        string strCom = "select * from [" + _sheet + "$" + _range + "]";
                        cn.Open();
                        OdbcCommand comm_mon = new OdbcCommand(strCom, cn);
                        OdbcDataAdapter da = new OdbcDataAdapter();
                        da.SelectCommand = comm_mon;
                        System.Data.DataTable dt = new System.Data.DataTable();
                        da.Fill(dt);
                        DataSaver ds = new DataSaver(login_, password_);
                        Warning = ds.SaveKSSStable(dt, Convert.ToInt32(tb_KSSS_firstRow.Text));
 
                        if (Warning != "")
                            MessageBox.Show(Warning);
                    }
                    else
                    {
                        MessageBox.Show("Укажите корректное количество строк!");
                    }
                }
                KSSStoContragentFill();
            }

            private void listBox_KSSS_SelectedIndexChanged(object sender, EventArgs e)
            {
                if (this.listBox_KSSS.DataSource != null)
                {
                    _sheet = this.listBox_KSSS.SelectedValue.ToString();
                    char[] charsToTrim = { '$', '#' };
                    _sheet = _sheet.Trim(charsToTrim);
                }
            }

            private void label18_Click(object sender, EventArgs e)
            {

            }

            private void tb_pay_PFMcode_TextChanged(object sender, EventArgs e)
            {

            }

            private void lBlockMessage_Click(object sender, EventArgs e)
            {

            }

            private void btn_del_actual_Click(object sender, EventArgs e)
            {
                int focusRow = ActualView.FocusedRowHandle;
                string idRow = "";
                idRow = ActualView.GetFocusedDataRow()["Idactual"].ToString();
                DataSaver ds = new DataSaver(login_, password_);
                lb_status.Text = ds.DeleteActualPayments(idRow);

                ActualPaymentsFill();
                try
                {
                    if (focusRow != ActualView.RowCount - 1)
                        ActualView.FocusedRowHandle = focusRow - 1;
                    else
                        ActualView.FocusedRowHandle = focusRow;
                }
                catch { }
                _connection.Close();
            }

            private void dateTimePicker8_ValueChanged(object sender, EventArgs e)
            {
                ActualPaymentsFill();
            }

            private void dateTimePicker9_ValueChanged(object sender, EventArgs e)
            {
                ActualPaymentsFill();
            }

            private void dateTimePicker10_ValueChanged(object sender, EventArgs e)
            {
                ContractsFill();
                dateTimePicker7.Value = dateTimePicker10.Value;
            }

            private void dateTimePicker7_ValueChanged(object sender, EventArgs e)
            {
                ContractsFill();
                dateTimePicker10.Value = dateTimePicker7.Value;
            }

            private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
            {
                PlansFill();
            }

            private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
            {
                PlansFill();
            }

            private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
            {
                LimitsFill();
            }

            private void textBox3_TextChanged(object sender, EventArgs e)
            {
                LimitsFill();
            }

            private void otch7_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
            {
                ReportMark reportForm = new ReportMark(_connection, TemplatePath, "MonthExpectedExecution");
                reportForm.Text = "Ожидаемое исполнение месяца";
                reportForm.Show();
            }

    }
}
