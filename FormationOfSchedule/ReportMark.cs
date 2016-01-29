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
    public partial class ReportMark : Form
    {

        SqlConnection _connection;
        string TemplatePath;
        string reportName;

        public ReportMark(SqlConnection connection, string _TemplatePath, string repName)
        {
            InitializeComponent();
            _connection = connection;
            TemplatePath = _TemplatePath;
            reportName = repName;
        }

        private void btn_Close_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private string CellName(int CellRow, int CellColumn)
        {
            return ((char)(64 + CellColumn) + CellRow.ToString());
        }

        private void btn_OK_Click(object sender, EventArgs e)
        {
            if (reportName == "MonthExpectedExecution")
                ReportMonthExpectedExecution();
            else if (reportName == "ReportMark")
            {
                // Путь к файлу шаблона
                string TemplateFileName = "ReportMark.xlt";
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
                int clWeightDifference = 21;		    // Разница по весу
                // Константы Excel
                int xlDown = -4121;
                // Закоментированы для того, чтобы не было Warning-ов
                int clShipmentNumber = 1;			// Номер отгрузки

                int DataRowNumber;
                int DataColumnNumber;

                string Comm = "SELECT * FROM udf_ReportMark (@data_)";
                SqlCommand ExelComm = new SqlCommand(Comm, _connection);

                ExelComm.Parameters.Add("@data_", SqlDbType.Date).Direction = ParameterDirection.Input;
                ExelComm.Parameters["@data_"].Value = dateTimePicker1.Value.Date;


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
                        if (ReportWorksheet.get_Range(CellName(CurrentRow, CurrentColumn), CellName(CurrentRow, CurrentColumn)).Text.ToString() == "#start2")
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

                    object[,] ReportDataArray = new object[ds.Tables[0].Rows.Count, 21];

                    for (DataRowNumber = 0; DataRowNumber < ds.Tables[0].Rows.Count; DataRowNumber++)
                    {
                        for (DataColumnNumber = 0; DataColumnNumber < 21; DataColumnNumber++)
                        {
                            ReportDataArray[DataRowNumber, DataColumnNumber] = ds.Tables[0].Rows[DataRowNumber][DataColumnNumber];
                        }
                    }
                    ReportWorksheet.get_Range(CellName(StartRow, clShipmentNumber), CellName(ds.Tables[0].Rows.Count + StartRow - 1, clWeightDifference)).Value2 = ReportDataArray;
                    // Отображение отчёта на экране



                    //////////////////////////////////////

                    Comm = "SELECT * FROM udf_ReportMark_title (@data_)";
                    SqlCommand ExelComm1 = new SqlCommand(Comm, _connection);

                    ExelComm1.Parameters.Add("@data_", SqlDbType.Date).Direction = ParameterDirection.Input;
                    ExelComm1.Parameters["@data_"].Value = dateTimePicker1.Value.Date;

                    DataSet ds1 = new DataSet();
                    try
                    {
                        if (_connection.State != System.Data.ConnectionState.Open)
                            _connection.Open();
                        ExelComm1.ExecuteNonQuery();

                        SqlDataAdapter new_dataAdapter = new SqlDataAdapter(ExelComm1);
                        new_dataAdapter.Fill(ds1);
                    }
                    catch { }

                    DataRowNumber = 0;
                    DataColumnNumber = 0;
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
                            if (ReportWorksheet.get_Range(CellName(CurrentRow, CurrentColumn), CellName(CurrentRow, CurrentColumn)).Text.ToString() == "#start1")
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
                        for (CurrentRow = StartRow; CurrentRow < ds1.Tables[0].Rows.Count + StartRow - 1; CurrentRow++)
                        {
                            ReportWorksheet.get_Range(CellName(CurrentRow, clWeightDifference), CellName(CurrentRow, clWeightDifference)).EntireRow.Insert(xlDown, true);
                        }

                        object[,] ReportDataArray1 = new object[ds1.Tables[0].Rows.Count, 21];

                        for (DataRowNumber = 0; DataRowNumber < ds1.Tables[0].Rows.Count; DataRowNumber++)
                        {
                            for (DataColumnNumber = 0; DataColumnNumber < 21; DataColumnNumber++)
                            {
                                ReportDataArray1[DataRowNumber, DataColumnNumber] = ds1.Tables[0].Rows[DataRowNumber][DataColumnNumber];
                            }
                        }
                        ReportWorksheet.get_Range(CellName(StartRow, clShipmentNumber), CellName(ds1.Tables[0].Rows.Count + StartRow - 1, clWeightDifference)).Value2 = ReportDataArray1;





                        ExcelApp.Visible = true;
                        this.Close();
                    }

                }

            }
        }

        private void ReportMonthExpectedExecution()
        {            
            SqlCommand comm = new SqlCommand("SELECT PFMName, PartnerType, FinPositionEPL, PlanSumm, FactSummOnDate, PlanSummOnDate, ExpectedMonthExec, AbsDeviation, RelativeDeviation FROM udf_get_MonthExpectedExecution(@date)", _connection);
            comm.CommandType = CommandType.Text;
            comm.Parameters.Add("@date", SqlDbType.Date).Direction = ParameterDirection.Input;
            comm.Parameters["@date"].Value = dateTimePicker1.Value.Date;

            try
            {
                if (_connection.State != System.Data.ConnectionState.Open)
                    _connection.Open();

                using (SqlDataReader reader = comm.ExecuteReader())
                {
                    //if (ExcelApp == null)
                    Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                    ExcelApp.Workbooks.Add(TemplatePath + "MonthExpectedExecution.xlt");
                    Microsoft.Office.Interop.Excel.Worksheet ReportWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelApp.ActiveWorkbook.Worksheets["First"];

                    Dictionary<string, string> labelDict = new Dictionary<string, string>();
                    labelDict.Add("#header", "Ожидаемое исполнение месяца на " + dateTimePicker1.Value.ToShortDateString());
                    labelDict.Add("#planmonth", "План поступлений/платежей на " + dateTimePicker1.Value.ToString("MM.yyyy"));
                    labelDict.Add("#factondate", "Факт " + "01." + dateTimePicker1.Value.ToString("MM") + " - " + dateTimePicker1.Value.ToString("dd.MM"));
                    labelDict.Add("#planondate", "План " + dateTimePicker1.Value.AddDays(1).ToString("dd.MM") + " - " + DateTime.DaysInMonth(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month) + "." + dateTimePicker1.Value.ToString("MM"));

                    string error;

                    if (!ReportExcelUtil.ExportData2Report(ReportWorksheet, reader, labelDict, out error))
                    {
                        ExcelApp.Quit();
                        MessageBox.Show(error, "Ошибка");
                    }
                    else
                    {
                        ExcelApp.Visible = true;
                        this.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }            
        }
    }
}