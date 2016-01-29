using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;

namespace FormationOfSchedule
{
    public partial class ReportLimits : Form
    {
        SqlConnection _connection;
        string TemplatePath;
        int _report_num;
        public ReportLimits(SqlConnection connection, int report_num, string _TemplatePath)
        {
            InitializeComponent();
            _connection = connection;
            _report_num = report_num;
            TemplatePath = _TemplatePath;
            if (report_num == 1)
            {
                label1.Text = "Выберите квартал";
                comboBox1.Visible = false;
                radioButton1.Visible = true;
                radioButton2.Visible = true;
                radioButton1.Checked = true;
            }
            else
            {
                label1.Text = "Выберите месяц";
                comboBox2.Visible = false;
                radioButton1.Visible = true;
                radioButton2.Visible = true;
                radioButton1.Checked = true;
            }

            tb_year.Text = DateTime.Now.Date.Year.ToString();
        }

        private void ReportLimits_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'forMonthLimits.Month' table. You can move, or remove it, as needed.
            this.monthTableAdapter.Fill(this.forMonthLimits.Month);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (_report_num == 1 && radioButton1.Checked)
            {
                // Путь к файлу шаблона
                string TemplateFileName = "KvartalLimits.xlt";
                TemplatePath = TemplatePath + TemplateFileName;
                string TemplateWorksheetName = "Отчет";

                // Константы для поиска стартовой ячейки вывода данных отчёта
                int MaxRowsToFindStart = 15;		// Просматриваемых строк
                int MaxColumnsToFindStart = 10;		// Просматриваемых столбцов

                // Переменные для перебора строк
                bool StartIsFind;					// Флаг найденного старта
                int CurrentRow;						// Текущая строка
                int CurrentColumn;					// Текущий столбец
                int StartRow;						// Стартовая строка
                int StartColumn;					// Стартовый столбец
                int clWeightDifference = 6;		    // Разница по весу
                // Константы Excel
                int xlDown = -4121;
                // Закоментированы для того, чтобы не было Warning-ов
                int clShipmentNumber = 1;			// Номер отгрузки

                int DataRowNumber;
                int DataColumnNumber;
                
                string Comm = "SELECT * FROM [udf_FS_Report_Climits] (@year, @kvartal)";
                SqlCommand ExelComm = new SqlCommand(Comm, _connection);

                ExelComm.Parameters.Add("@year", SqlDbType.VarChar).Direction = ParameterDirection.Input;
                ExelComm.Parameters["@year"].Value = tb_year.Text;

                ExelComm.Parameters.Add("@kvartal", SqlDbType.VarChar).Direction = ParameterDirection.Input;
                ExelComm.Parameters["@kvartal"].Value = comboBox2.Text.Remove(1).ToString();

                DataSet ds = new DataSet();
                try
                {
                    if (_connection.State != System.Data.ConnectionState.Open)
                        _connection.Open();
                    ExelComm.ExecuteNonQuery();

                    SqlDataAdapter new_dataAdapter = new SqlDataAdapter(ExelComm);
                    new_dataAdapter.Fill(ds);
                }
                catch { }

                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                // Создание книги по шаблону
                ExcelApp.Workbooks.Add(TemplatePath);
                // Получение ссылки на лист отчёта
                Microsoft.Office.Interop.Excel.Worksheet ReportWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelApp.ActiveWorkbook.Worksheets[TemplateWorksheetName];
                // Поиск ячейки с которой должен начатся вывод данных
                CurrentRow = 2;
                CurrentColumn = 1;
                StartIsFind = false;
                StartRow = 1;
                StartColumn = 1;

                while (!StartIsFind && CurrentColumn <= MaxColumnsToFindStart)
                {
                    // Поиск стартовой ячейки
                    // Поиск в текущем столбце
                    while (!StartIsFind && CurrentRow <= MaxRowsToFindStart)
                    {
                        // Проверка текущей ячейки
                        if (ReportWorksheet.get_Range(CellName(CurrentRow, CurrentColumn), CellName(CurrentRow, CurrentColumn)).Text.ToString() == "#start")
                        {
                            StartIsFind = true;
                            StartRow = CurrentRow;
                            StartColumn = CurrentColumn;
                        }
                        // Переход к следующей ячейке
                        CurrentRow++;
                    }
                    // Переход к следующему столбцу
                    CurrentColumn++;
                    CurrentRow = 1;
                }

                // Вывод данных отчёта начиная со стартовой позиции
                if (StartIsFind)
                {
                    // Вывод данных в шаблон
                    // Переход к начальной позиции вывода данных
                    CurrentRow = StartRow;
                    // Вывод строк из таблицы
                    // Добавление нужного количества строк
                    for (CurrentRow = StartRow; CurrentRow < ds.Tables[0].Rows.Count + StartRow - 1; CurrentRow++)
                    {
                        ReportWorksheet.get_Range(CellName(CurrentRow, clWeightDifference), CellName(CurrentRow, clWeightDifference)).EntireRow.Insert(xlDown, true);
                    }

                    object[,] ReportDataArray = new object[ds.Tables[0].Rows.Count, 6];

                    for (DataRowNumber = 0; DataRowNumber < ds.Tables[0].Rows.Count; DataRowNumber++)
                    {
                        for (DataColumnNumber = 0; DataColumnNumber < 6; DataColumnNumber++)
                        {
                            ReportDataArray[DataRowNumber, DataColumnNumber] = ds.Tables[0].Rows[DataRowNumber][DataColumnNumber];
                        }
                    }
                    ReportWorksheet.get_Range(CellName(StartRow, clShipmentNumber), CellName(ds.Tables[0].Rows.Count + StartRow - 1, clWeightDifference)).Value2 = ReportDataArray;
                    ReportWorksheet.get_Range(CellName(3, 3), CellName(3, 3)).Value2 = "за  " + comboBox2.Text.Remove(1).ToString() + "  квартал  " + tb_year.Text + "  года";
                    // Отображение отчёта на экране
                    ExcelApp.Visible = true;
                    this.Close();
                }
            }


            if (_report_num == 1 && radioButton2.Checked)
            {
                // Путь к файлу шаблона
                string TemplateFileName = "KvartalLimits Plan.xlt";
                TemplatePath = TemplatePath + TemplateFileName;
                string TemplateWorksheetName = "Отчет";

                // Константы для поиска стартовой ячейки вывода данных отчёта
                int MaxRowsToFindStart = 15;		// Просматриваемых строк
                int MaxColumnsToFindStart = 10;		// Просматриваемых столбцов

                // Переменные для перебора строк
                bool StartIsFind;					// Флаг найденного старта
                int CurrentRow;						// Текущая строка
                int CurrentColumn;					// Текущий столбец
                int StartRow;						// Стартовая строка
                int StartColumn;					// Стартовый столбец
                int clWeightDifference = 6;		    // Разница по весу
                // Константы Excel
                int xlDown = -4121;
                // Закоментированы для того, чтобы не было Warning-ов
                int clShipmentNumber = 1;			// Номер отгрузки

                int DataRowNumber;
                int DataColumnNumber;

                string Comm = "SELECT * FROM [udf_FS_Report_Climits_Plan] (@year, @kvartal)";
                SqlCommand ExelComm = new SqlCommand(Comm, _connection);

                ExelComm.Parameters.Add("@year", SqlDbType.VarChar).Direction = ParameterDirection.Input;
                ExelComm.Parameters["@year"].Value = tb_year.Text;

                ExelComm.Parameters.Add("@kvartal", SqlDbType.VarChar).Direction = ParameterDirection.Input;
                ExelComm.Parameters["@kvartal"].Value = comboBox2.Text.Remove(1).ToString();

                DataSet ds = new DataSet();
                try
                {
                    if (_connection.State != System.Data.ConnectionState.Open)
                        _connection.Open();
                    ExelComm.ExecuteNonQuery();

                    SqlDataAdapter new_dataAdapter = new SqlDataAdapter(ExelComm);
                    new_dataAdapter.Fill(ds);
                }
                catch { }

                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                // Создание книги по шаблону
                ExcelApp.Workbooks.Add(TemplatePath);
                // Получение ссылки на лист отчёта
                Microsoft.Office.Interop.Excel.Worksheet ReportWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelApp.ActiveWorkbook.Worksheets[TemplateWorksheetName];
                // Поиск ячейки с которой должен начатся вывод данных
                CurrentRow = 2;
                CurrentColumn = 1;
                StartIsFind = false;
                StartRow = 1;
                StartColumn = 1;

                while (!StartIsFind && CurrentColumn <= MaxColumnsToFindStart)
                {
                    // Поиск стартовой ячейки
                    // Поиск в текущем столбце
                    while (!StartIsFind && CurrentRow <= MaxRowsToFindStart)
                    {
                        // Проверка текущей ячейки
                        if (ReportWorksheet.get_Range(CellName(CurrentRow, CurrentColumn), CellName(CurrentRow, CurrentColumn)).Text.ToString() == "#start")
                        {
                            StartIsFind = true;
                            StartRow = CurrentRow;
                            StartColumn = CurrentColumn;
                        }
                        // Переход к следующей ячейке
                        CurrentRow++;
                    }
                    // Переход к следующему столбцу
                    CurrentColumn++;
                    CurrentRow = 1;
                }

                // Вывод данных отчёта начиная со стартовой позиции
                if (StartIsFind)
                {
                    // Вывод данных в шаблон
                    // Переход к начальной позиции вывода данных
                    CurrentRow = StartRow;
                    // Вывод строк из таблицы
                    // Добавление нужного количества строк
                    for (CurrentRow = StartRow; CurrentRow < ds.Tables[0].Rows.Count + StartRow - 1; CurrentRow++)
                    {
                        ReportWorksheet.get_Range(CellName(CurrentRow, clWeightDifference), CellName(CurrentRow, clWeightDifference)).EntireRow.Insert(xlDown, true);
                    }

                    object[,] ReportDataArray = new object[ds.Tables[0].Rows.Count, 6];

                    for (DataRowNumber = 0; DataRowNumber < ds.Tables[0].Rows.Count; DataRowNumber++)
                    {
                        for (DataColumnNumber = 0; DataColumnNumber < 6; DataColumnNumber++)
                        {
                            ReportDataArray[DataRowNumber, DataColumnNumber] = ds.Tables[0].Rows[DataRowNumber][DataColumnNumber];
                        }
                    }
                    ReportWorksheet.get_Range(CellName(StartRow, clShipmentNumber), CellName(ds.Tables[0].Rows.Count + StartRow - 1, clWeightDifference)).Value2 = ReportDataArray;
                    ReportWorksheet.get_Range(CellName(3, 3), CellName(3, 3)).Value2 = "за  " + comboBox2.Text.Remove(1).ToString() + "  квартал  " + tb_year.Text + "  года";
                    // Отображение отчёта на экране
                    ExcelApp.Visible = true;
                    this.Close();
                }

            }


            if (_report_num == 2 && radioButton1.Checked)
            {
                // Путь к файлу шаблона
                string TemplateFileName = "MonthLimits.xlt";
                TemplatePath = TemplatePath + TemplateFileName;
                string TemplateWorksheetName = "Отчет";

                // Константы для поиска стартовой ячейки вывода данных отчёта
                int MaxRowsToFindStart = 15;		// Просматриваемых строк
                int MaxColumnsToFindStart = 10;		// Просматриваемых столбцов

                // Переменные для перебора строк
                bool StartIsFind;					// Флаг найденного старта
                int CurrentRow;						// Текущая строка
                int CurrentColumn;					// Текущий столбец
                int StartRow;						// Стартовая строка
                int StartColumn;					// Стартовый столбец
                int clWeightDifference = 9;		    // Разница по весу
                // Константы Excel
                int xlDown = -4121;
                // Закоментированы для того, чтобы не было Warning-ов
                int clShipmentNumber = 1;			// Номер отгрузки

                int DataRowNumber;
                int DataColumnNumber;

                string Comm = "SELECT * FROM [udf_FS_Report_MLimits] (@month ,@year)";
                SqlCommand ExelComm = new SqlCommand(Comm, _connection);

                ExelComm.Parameters.Add("@year", SqlDbType.VarChar).Direction = ParameterDirection.Input;
                ExelComm.Parameters["@year"].Value = tb_year.Text;

                ExelComm.Parameters.Add("@month", SqlDbType.VarChar).Direction = ParameterDirection.Input;
                ExelComm.Parameters["@month"].Value = comboBox1.SelectedValue.ToString();

                DataSet ds = new DataSet();
                try
                {
                    if (_connection.State != System.Data.ConnectionState.Open)
                        _connection.Open();
                    ExelComm.ExecuteNonQuery();

                    SqlDataAdapter new_dataAdapter = new SqlDataAdapter(ExelComm);
                    new_dataAdapter.Fill(ds);
                }
                catch { }

                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                // Создание книги по шаблону
                ExcelApp.Workbooks.Add(TemplatePath);
                // Получение ссылки на лист отчёта
                Microsoft.Office.Interop.Excel.Worksheet ReportWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelApp.ActiveWorkbook.Worksheets[TemplateWorksheetName];
                // Поиск ячейки с которой должен начатся вывод данных
                CurrentRow = 2;
                CurrentColumn = 1;
                StartIsFind = false;
                StartRow = 1;
                StartColumn = 1;

                while (!StartIsFind && CurrentColumn <= MaxColumnsToFindStart)
                {
                    // Поиск стартовой ячейки
                    // Поиск в текущем столбце
                    while (!StartIsFind && CurrentRow <= MaxRowsToFindStart)
                    {
                        // Проверка текущей ячейки
                        if (ReportWorksheet.get_Range(CellName(CurrentRow, CurrentColumn), CellName(CurrentRow, CurrentColumn)).Text.ToString() == "#start")
                        {
                            StartIsFind = true;
                            StartRow = CurrentRow;
                            StartColumn = CurrentColumn;
                        }
                        // Переход к следующей ячейке
                        CurrentRow++;
                    }
                    // Переход к следующему столбцу
                    CurrentColumn++;
                    CurrentRow = 1;
                }

                // Вывод данных отчёта начиная со стартовой позиции
                if (StartIsFind)
                {
                    // Вывод данных в шаблон
                    // Переход к начальной позиции вывода данных
                    CurrentRow = StartRow;
                    // Вывод строк из таблицы
                    // Добавление нужного количества строк
                    for (CurrentRow = StartRow; CurrentRow < ds.Tables[0].Rows.Count + StartRow - 1; CurrentRow++)
                    {
                        ReportWorksheet.get_Range(CellName(CurrentRow, clWeightDifference), CellName(CurrentRow, clWeightDifference)).EntireRow.Insert(xlDown, true);
                    }

                    object[,] ReportDataArray = new object[ds.Tables[0].Rows.Count, 9];

                    for (DataRowNumber = 0; DataRowNumber < ds.Tables[0].Rows.Count; DataRowNumber++)
                    {
                        for (DataColumnNumber = 0; DataColumnNumber < 9; DataColumnNumber++)
                        {
                            ReportDataArray[DataRowNumber, DataColumnNumber] = ds.Tables[0].Rows[DataRowNumber][DataColumnNumber];
                        }
                    }
                    ReportWorksheet.get_Range(CellName(StartRow, clShipmentNumber), CellName(ds.Tables[0].Rows.Count + StartRow - 1, clWeightDifference)).Value2 = ReportDataArray;
                    ReportWorksheet.get_Range(CellName(3, 4), CellName(3, 4)).Value2 = "за  " + comboBox1.Text.ToString() + "  месяц  " + tb_year.Text + "  года";
                    // Отображение отчёта на экране
                    ExcelApp.Visible = true;
                    this.Close();
                }
            }
            if (_report_num == 2 && radioButton2.Checked)
            {
                // Путь к файлу шаблона
                string TemplateFileName = "MonthLimits Plan.xlt";
                TemplatePath = TemplatePath + TemplateFileName;
                string TemplateWorksheetName = "Отчет";

                // Константы для поиска стартовой ячейки вывода данных отчёта
                int MaxRowsToFindStart = 15;		// Просматриваемых строк
                int MaxColumnsToFindStart = 10;		// Просматриваемых столбцов

                // Переменные для перебора строк
                bool StartIsFind;					// Флаг найденного старта
                int CurrentRow;						// Текущая строка
                int CurrentColumn;					// Текущий столбец
                int StartRow;						// Стартовая строка
                int StartColumn;					// Стартовый столбец
                int clWeightDifference = 9;		    // Разница по весу
                // Константы Excel
                int xlDown = -4121;
                // Закоментированы для того, чтобы не было Warning-ов
                int clShipmentNumber = 1;			// Номер отгрузки

                int DataRowNumber;
                int DataColumnNumber;

                string Comm = "SELECT * FROM [udf_FS_Report_MLimits_Plan] (@month ,@year)";
                SqlCommand ExelComm = new SqlCommand(Comm, _connection);

                ExelComm.Parameters.Add("@year", SqlDbType.VarChar).Direction = ParameterDirection.Input;
                ExelComm.Parameters["@year"].Value = tb_year.Text;

                ExelComm.Parameters.Add("@month", SqlDbType.VarChar).Direction = ParameterDirection.Input;
                ExelComm.Parameters["@month"].Value = comboBox1.SelectedValue.ToString();

                DataSet ds = new DataSet();
                try
                {
                    if (_connection.State != System.Data.ConnectionState.Open)
                        _connection.Open();
                    ExelComm.ExecuteNonQuery();

                    SqlDataAdapter new_dataAdapter = new SqlDataAdapter(ExelComm);
                    new_dataAdapter.Fill(ds);
                }
                catch { }

                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                // Создание книги по шаблону
                ExcelApp.Workbooks.Add(TemplatePath);
                // Получение ссылки на лист отчёта
                Microsoft.Office.Interop.Excel.Worksheet ReportWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelApp.ActiveWorkbook.Worksheets[TemplateWorksheetName];
                // Поиск ячейки с которой должен начатся вывод данных
                CurrentRow = 2;
                CurrentColumn = 1;
                StartIsFind = false;
                StartRow = 1;
                StartColumn = 1;

                while (!StartIsFind && CurrentColumn <= MaxColumnsToFindStart)
                {
                    // Поиск стартовой ячейки
                    // Поиск в текущем столбце
                    while (!StartIsFind && CurrentRow <= MaxRowsToFindStart)
                    {
                        // Проверка текущей ячейки
                        if (ReportWorksheet.get_Range(CellName(CurrentRow, CurrentColumn), CellName(CurrentRow, CurrentColumn)).Text.ToString() == "#start")
                        {
                            StartIsFind = true;
                            StartRow = CurrentRow;
                            StartColumn = CurrentColumn;
                        }
                        // Переход к следующей ячейке
                        CurrentRow++;
                    }
                    // Переход к следующему столбцу
                    CurrentColumn++;
                    CurrentRow = 1;
                }

                // Вывод данных отчёта начиная со стартовой позиции
                if (StartIsFind)
                {
                    // Вывод данных в шаблон
                    // Переход к начальной позиции вывода данных
                    CurrentRow = StartRow;
                    // Вывод строк из таблицы
                    // Добавление нужного количества строк
                    for (CurrentRow = StartRow; CurrentRow < ds.Tables[0].Rows.Count + StartRow - 1; CurrentRow++)
                    {
                        ReportWorksheet.get_Range(CellName(CurrentRow, clWeightDifference), CellName(CurrentRow, clWeightDifference)).EntireRow.Insert(xlDown, true);
                    }

                    object[,] ReportDataArray = new object[ds.Tables[0].Rows.Count, 9];

                    for (DataRowNumber = 0; DataRowNumber < ds.Tables[0].Rows.Count; DataRowNumber++)
                    {
                        for (DataColumnNumber = 0; DataColumnNumber < 9; DataColumnNumber++)
                        {
                            ReportDataArray[DataRowNumber, DataColumnNumber] = ds.Tables[0].Rows[DataRowNumber][DataColumnNumber];
                        }
                    }
                    ReportWorksheet.get_Range(CellName(StartRow, clShipmentNumber), CellName(ds.Tables[0].Rows.Count + StartRow - 1, clWeightDifference)).Value2 = ReportDataArray;
                    ReportWorksheet.get_Range(CellName(3, 4), CellName(3, 4)).Value2 = "за  " + comboBox1.Text.ToString() + "  месяц  " + tb_year.Text + "  года";
                    // Отображение отчёта на экране
                    ExcelApp.Visible = true;
                    this.Close();
                }
            }
        }

        private string CellName(int CellRow, int CellColumn)
        {
            return ((char)(64 + CellColumn) + CellRow.ToString());
        }
    }
}
