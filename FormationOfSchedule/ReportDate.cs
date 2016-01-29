using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;

namespace FormationOfSchedule
{
    public partial class ReportDate : Form
    {
        SqlConnection _connection;
        int _otch_num;
        string TemplatePath;

        Microsoft.Office.Interop.Excel.Application ExcelApp = null;

        public ReportDate(SqlConnection connection, int otch_num, string _TemplatePath)
        {
            InitializeComponent();
            _connection = connection;
            _otch_num = otch_num;
            TemplatePath = _TemplatePath;
        }        

        private void btn_OK_Click(object sender, EventArgs e)
        {
            // Путь к файлу шаблона
            string TemplateFileName = "ReportIncludePayments.xlt";
            //TemplatePath = TemplatePath + TemplateFileName;
            //string TemplatePath = TemplateFileName;
            string TemplateWorksheetName = "Отчет";
            string TemplateWorksheetName2 = "Отчет по декадам";
            // Константы для поиска стартовой ячейки вывода данных отчёта
            int MaxRowsToFindStart = 18;		// Просматриваемых строк
            int MaxColumnsToFindStart = 10;		// Просматриваемых столбцов

            // Переменные для перебора строк
            bool StartIsFind;					// Флаг найденного старта
            int CurrentRow;						// Текущая строка
            int CurrentColumn;					// Текущий столбец
            int StartRow;						// Стартовая строка
            int StartColumn;					// Стартовый столбец
            int clWeightDifference = 9;		// Разница по весу
            bool StartIsFind2;					// Флаг найденного старта
            int CurrentRow2;						// Текущая строка
            int CurrentColumn2;					// Текущий столбец
            int StartRow2;						// Стартовая строка
            int StartColumn2;					// Стартовый столбец
            int clWeightDifference2 = 6;		// Разница по весу
            // Константы Excel
            int xlDown = -4121;
            // Закоментированы для того, чтобы не было Warning-ов
            int clShipmentNumber = 1;			// Номер отгрузки

            int DataRowNumber;
            int DataColumnNumber;
            //string date1 = dateTimePicker2.Value.ToString();
            // string date2;
            // string pfm = _information[1];
            string Comm = "";
            string Comm2 = "";
            if(_otch_num == 2)
            Comm = "SELECT * FROM [udf_Report_Date] (@startDate,@endDate,2)";
            else
                Comm = "SELECT * FROM [udf_Report_Date] (@startDate,@endDate,1)";
            Comm2 = "SELECT * FROM [udf_ReportDecade] (@data)";

            SqlCommand ExelComm = new SqlCommand(Comm, _connection);

            ExelComm.Parameters.Add("@startDate", SqlDbType.Date).Direction = ParameterDirection.Input;
            ExelComm.Parameters["@startDate"].Value = dateTimePicker1.Value.Date;

            ExelComm.Parameters.Add("@endDate", SqlDbType.Date).Direction = ParameterDirection.Input;
            ExelComm.Parameters["@endDate"].Value = dateTimePicker2.Value.Date;

            DataSet ds = new DataSet();

            SqlCommand ExelComm2 = new SqlCommand(Comm2, _connection);
            ExelComm2.Parameters.Add("@data", SqlDbType.Date).Direction = ParameterDirection.Input;
            ExelComm2.Parameters["@data"].Value = dateTimePicker1.Value.Date;
            DataSet ds2 = new DataSet();


            try
            {
                if (_connection.State != System.Data.ConnectionState.Open)
                    _connection.Open();
                ExelComm.ExecuteNonQuery();

                SqlDataAdapter new_dataAdapter = new SqlDataAdapter(ExelComm);
                new_dataAdapter.Fill(ds);
            }
            catch { }

            try
            {
                if (_connection.State != System.Data.ConnectionState.Open)
                    _connection.Open();
                ExelComm2.ExecuteNonQuery();

                SqlDataAdapter new_dataAdapter2 = new SqlDataAdapter(ExelComm2);
                new_dataAdapter2.Fill(ds2);
            }
            catch { }

            if (ExcelApp == null)
            /*Microsoft.Office.Interop.Excel.Application*/ ExcelApp = new Microsoft.Office.Interop.Excel.Application();

            // Создание книги по шаблону
            ExcelApp.Workbooks.Add(TemplatePath + TemplateFileName);
            // Получение ссылки на лист отчёта
            Microsoft.Office.Interop.Excel.Worksheet ReportWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelApp.ActiveWorkbook.Worksheets[TemplateWorksheetName];
            Microsoft.Office.Interop.Excel.Worksheet ReportWorksheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ExcelApp.ActiveWorkbook.Worksheets[TemplateWorksheetName2];
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

            //Для второго отчета
            // Поиск ячейки с которой должен начатся вывод данных
            CurrentRow2 = 2;
            CurrentColumn2 = 1;
            StartIsFind2 = false;
            StartRow2 = 1;
            StartColumn2 = 1;

            while (!StartIsFind2 && CurrentColumn2 <= MaxColumnsToFindStart)
            {
                // Поиск стартовой ячейки
                // Поиск в текущем столбце
                while (!StartIsFind2 && CurrentRow2 <= MaxRowsToFindStart)
                {
                    // Проверка текущей ячейки
                    if (ReportWorksheet2.get_Range(CellName(CurrentRow2, CurrentColumn2), CellName(CurrentRow2, CurrentColumn2)).Text.ToString() == "#start")
                    {
                        StartIsFind2 = true;
                        StartRow2 = CurrentRow2;
                        StartColumn2 = CurrentColumn2;
                    }
                    // Переход к следующей ячейке
                    CurrentRow2++;
                }
                // Переход к следующему столбцу
                CurrentColumn2++;
                CurrentRow2 = 1;
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
                
                if(_otch_num == 2)
                ExcelApp.GetType().InvokeMember("Run", System.Reflection.BindingFlags.Default | System.Reflection.BindingFlags.InvokeMethod, null, ExcelApp, new object[] { "Svodnaia" });

            }
            //Для второго отчета
            if (StartIsFind2)
            {
                // Вывод данных в шаблон
                // Переход к начальной позиции вывода данных
                CurrentRow2 = StartRow2;
                // Вывод строк из таблицы
                // Добавление нужного количества строк
                for (CurrentRow2 = StartRow2; CurrentRow2 < ds2.Tables[0].Rows.Count + StartRow2 - 1; CurrentRow2++)
                {
                    ReportWorksheet2.get_Range(CellName(CurrentRow2, clWeightDifference2), CellName(CurrentRow2, clWeightDifference2)).EntireRow.Insert(xlDown, true);
                }

                object[,] ReportDataArray2 = new object[ds2.Tables[0].Rows.Count, 6];

                for (DataRowNumber = 0; DataRowNumber < ds2.Tables[0].Rows.Count; DataRowNumber++)
                {
                    for (DataColumnNumber = 0; DataColumnNumber < 6; DataColumnNumber++)
                    {
                        ReportDataArray2[DataRowNumber, DataColumnNumber] = ds2.Tables[0].Rows[DataRowNumber][DataColumnNumber];
                    }
                }


                ReportWorksheet2.get_Range(CellName(StartRow2, clShipmentNumber), CellName(ds2.Tables[0].Rows.Count + StartRow2 - 1, clWeightDifference2)).Value2 = ReportDataArray2;
                ReportWorksheet2.get_Range(CellName(2, 3), CellName(2, 3)).Value2 = dateTimePicker1.Value.Date.Day.ToString() + "."
                                                                                    + dateTimePicker1.Value.Date.Month.ToString() + "." 
                                                                                    + dateTimePicker1.Value.Date.Year.ToString() + " - "
                                                                                    + dateTimePicker2.Value.Date.Day.ToString() + "."
                                                                                    + dateTimePicker2.Value.Date.Month.ToString() + "."
                                                                                    + dateTimePicker2.Value.Date.Year.ToString();
                }

            if (StartIsFind || StartIsFind2)
            {
                // Отображение отчёта на экране
                ExcelApp.Visible = true;

                //сохранение результатов отчета по декадам для отчета средней оценки точности
                if (MessageBox.Show("Сохранить результаты отчета по декадам?", "Сообщение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                {
                    if (_connection != null)
                    {
                        DataSaver saver = new DataSaver(_connection);

                        string mess = saver.InsertPaymentsPlanSum(dateTimePicker1.Value.Date);

                        MessageBox.Show(mess, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }

            this.Close();
            
        }

        private string CellName(int CellRow, int CellColumn)
        {
            return ((char)(64 + CellColumn) + CellRow.ToString());
        }

        private void ReportDate_Load(object sender, EventArgs e)
        {
            SqlCommand comm = new SqlCommand("select period from udf_get_PaymentsPlanSumPeriods()", _connection);
            comm.CommandType = CommandType.Text;

            try
            {
                if (_connection.State != System.Data.ConnectionState.Open)
                    _connection.Open();

                using (var reader = comm.ExecuteReader())
                {
                    List<string> periods = reader.Cast<IDataRecord>().Select(dr => dr.GetDateTime(0).ToShortDateString()).ToList();
                    reader.Close();
                    listBoxPeriods.DataSource = periods;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void listBoxPeriods_SelectedValueChanged(object sender, EventArgs e)
        {
            DateTime period;
            gridControl1.DataSource = null;

            if (DateTime.TryParse(listBoxPeriods.SelectedValue.ToString(), out period))
            {
                SqlCommand comm = new SqlCommand("select PFMName, PartnerType, FinPositionEPL, Summ, DateStart,	DateEnd from udf_get_PaymentsPlanSum(@date)", _connection);
                comm.CommandType = CommandType.Text;

                comm.Parameters.Add("@date", SqlDbType.Date).Direction = ParameterDirection.Input;
                comm.Parameters["@date"].Value = period;

                try
                {
                    if (_connection.State != System.Data.ConnectionState.Open)
                        _connection.Open();

                    using (var reader = comm.ExecuteReader())
                    {
                        var paymentsSum = reader.Cast<IDataRecord>().Select(dr => new
                        {
                            PFMName = dr.GetString(0),
                            PartnerType = dr.GetString(1),
                            FinPositionEPL = dr.GetString(2),
                            Summ = dr.GetValue(3),
                            DateStart = dr.GetDateTime(4),
                            DateEnd = dr.GetDateTime(5)
                        }).ToList();

                        gridControl1.DataSource = paymentsSum;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btn_ShowDetail_Click(object sender, EventArgs e)
        {
            DateTime period;

            if (listBoxPeriods.SelectedValue == null)
            {
                MessageBox.Show("Не выбран период.", "Ошибка");
                return;
            }

            if (DateTime.TryParse(listBoxPeriods.SelectedValue.ToString(), out period))
            {
                SqlCommand comm = new SqlCommand("SELECT PFMName, PartnerType, FinPositionEPL, Summ, DatePay, DateStart, DateEnd FROM udf_get_PaymentsPlanSumDetail(@date)", _connection);
                comm.CommandType = CommandType.Text;

                comm.Parameters.Add("@date", SqlDbType.Date).Direction = ParameterDirection.Input;
                comm.Parameters["@date"].Value = period;

                try
                {
                    if (_connection.State != System.Data.ConnectionState.Open)
                        _connection.Open();

                    using (SqlDataReader reader = comm.ExecuteReader())
                    {
                        if (ExcelApp == null)
                            /*Microsoft.Office.Interop.Excel.Application*/ ExcelApp = new Microsoft.Office.Interop.Excel.Application();                        
                        ExcelApp.Workbooks.Add(TemplatePath + "ReportArchivePaymentSumDetail.xlt");                        
                        Microsoft.Office.Interop.Excel.Worksheet ReportWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelApp.ActiveWorkbook.Worksheets["First"];
                        
                        Dictionary<string, string> labelDict = new Dictionary<string, string>();
                        labelDict.Add("#header", "Отчетная форма из архивной базы данных на " + listBoxPeriods.SelectedValue);

                        string error;

                        if (!ReportExcelUtil.ExportData2Report(ReportWorksheet, reader, labelDict, out error))
                        {
                            ExcelApp.Quit();
                            MessageBox.Show(error, "Ошибка");
                        }
                        else                        
                            ExcelApp.Visible = true;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }            
        }        
    }

    public static class ReportExcelUtil
    {
        public static bool ExportData2Report(Microsoft.Office.Interop.Excel.Worksheet reportWorksheet, SqlDataReader reader, Dictionary<string, string> labelDict, out string errMessage)
        {
            foreach (var item in labelDict)
            {
                var label = reportWorksheet.Cells.Find(item.Key);
                if (label != null)                
                    label.Value2 = item.Value;                
            }
            //нахождение стартовой ячейки с заданной меткой
            var start = reportWorksheet.Cells.Find("#start");

            if (start == null)
            {
                errMessage = "Не найдена стартовая метка для вывода данных.";                
                reader.Dispose();
                return false;
            }
            start.Value2 = null;

            try
            {
                int row = start.Row, col = start.Column;
                int fieldCount = reader.FieldCount;
                object[] valArray = new object[fieldCount];

                Microsoft.Office.Interop.Excel.Range startCell;
                Microsoft.Office.Interop.Excel.Range endCell;
                Microsoft.Office.Interop.Excel.Range excelCells;

                while (reader.Read())
                {
                    startCell = (Microsoft.Office.Interop.Excel.Range)reportWorksheet.Cells[row, col];
                    endCell = (Microsoft.Office.Interop.Excel.Range)reportWorksheet.Cells[row, col + fieldCount - 1];
                    excelCells = reportWorksheet.Range[startCell, endCell];

                    for (int i = 0; i < fieldCount; i++)
                        valArray[i] = reader.GetValue(i);

                    excelCells.Value2 = valArray;
                    reportWorksheet.Rows[row + 1].EntireRow.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, true);
                    row++;
                }
                ((Microsoft.Office.Interop.Excel.Range)reportWorksheet.Rows[row]).EntireRow.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
            }
            catch (Exception ex)
            {
                errMessage = ex.Message;
                reader.Dispose();
                return false;
            }

            errMessage = string.Empty;
            return true;
        }
    }
}
