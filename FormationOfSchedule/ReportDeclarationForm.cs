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
    public partial class ReportDeclarationForm : Form
    {
        SqlConnection _connection;
        string[] _information;
        string TemplatePath;
        public ReportDeclarationForm(SqlConnection connection, string[] information, string _TemplatePath)
        {
            InitializeComponent();
            _connection = connection;
            _information = information;
            TemplatePath = _TemplatePath;
        }

        private void btn_Close_Otchet_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_Otchet_Click(object sender, EventArgs e)
        {
            // Путь к файлу шаблона
            string TemplateFileName = "Zaiavlenie.xlt";
             TemplatePath = TemplatePath + TemplateFileName;
            //string TemplatePath = TemplateFileName;
            string TemplateWorksheetName = "Отчет";

            // Константы для поиска стартовой ячейки вывода данных отчёта
            int MaxRowsToFindStart = 18;		// Просматриваемых строк
            int MaxColumnsToFindStart = 10;		// Просматриваемых столбцов

            // Переменные для перебора строк
            bool StartIsFind;					// Флаг найденного старта
            int CurrentRow;						// Текущая строка
            int CurrentColumn;					// Текущий столбец
            int StartRow;						// Стартовая строка
            int StartColumn;					// Стартовый столбец
            int clWeightDifference = 14;		// Разница по весу
            // Константы Excel
            int xlDown = -4121;
            // Закоментированы для того, чтобы не было Warning-ов
            int clShipmentNumber = 1;			// Номер отгрузки

            int DataRowNumber;
            int DataColumnNumber;
            //string date1 = dateTimePicker2.Value.ToString();
           // string date2;
           // string pfm = _information[1];

            string Comm = "SELECT * FROM [udf_FS_Report_Declaration] (@date1, @date2 , @pfm ,'Post') UNION ALL " +
                            " SELECT * FROM [udf_FS_Report_Declaration] (@date1 , @date2, @pfm ,'Vip') UNION ALL " +
                            " SELECT * FROM [udf_FS_Report_Declaration] (@date1, @date2, @pfm ,'Proch')";
           
            SqlCommand ExelComm = new SqlCommand(Comm, _connection);
            ExelComm.Parameters.Add("@date1", SqlDbType.DateTime).Direction = ParameterDirection.Input;
            ExelComm.Parameters["@date1"].Value = dateTimePicker2.Value.Date;

            ExelComm.Parameters.Add("@date2", SqlDbType.DateTime).Direction = ParameterDirection.Input;
            ExelComm.Parameters["@date2"].Value = dateTimePicker1.Value.Date;

            ExelComm.Parameters.Add("@pfm", SqlDbType.VarChar).Direction = ParameterDirection.Input;
            ExelComm.Parameters["@pfm"].Value = _information[1];

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
            CurrentRow = 4;
            CurrentColumn = 1;
            StartIsFind = false;
            StartRow = 1;
            StartColumn = 1;
            //ExcelApp.GetType().InvokeMember("Run", System.Reflection.BindingFlags.Default | System.Reflection.BindingFlags.InvokeMethod, null, ExcelApp, new object[] { }); 


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

                object[,] ReportDataArray = new object[ds.Tables[0].Rows.Count, 20];

                for (DataRowNumber = 0; DataRowNumber < ds.Tables[0].Rows.Count; DataRowNumber++)
                {
                    for (DataColumnNumber = 0; DataColumnNumber < 14; DataColumnNumber++)
                    {
                        ReportDataArray[DataRowNumber, DataColumnNumber] = ds.Tables[0].Rows[DataRowNumber][DataColumnNumber];
                    }
                }
                ReportWorksheet.get_Range(CellName(StartRow, clShipmentNumber), CellName(ds.Tables[0].Rows.Count + StartRow - 1, clWeightDifference)).Value2 = ReportDataArray;
                ReportWorksheet.get_Range(CellName(3, 3), CellName(3, 3)).Value2 = dateTimePicker2.Value.Date.Day + "."+ 
                    dateTimePicker2.Value.Date.Month + "." + dateTimePicker2.Value.Date.Year + "  -  " + 
                    dateTimePicker1.Value.Date.Day + "." + dateTimePicker1.Value.Date.Month + "." + 
                    dateTimePicker1.Value.Date.Year;

                string Comm2 = "SELECT PFMname FROM Users WHERE Users.PFMcode = '" + _information[1].ToString() + "'";

                SqlCommand ExelComm2 = new SqlCommand(Comm2, _connection);
                /*ExelComm2.Parameters.Add("@pfm", SqlDbType.VarChar).Direction = ParameterDirection.Input;
                ExelComm2.Parameters["@pfm"].Value = _information[1];
                */
                DataSet ds2 = new DataSet();
                object PfmName = "";
                try
                {
                    if (_connection.State != System.Data.ConnectionState.Open)
                        _connection.Open();
                    PfmName = ExelComm2.ExecuteScalar();

                    SqlDataAdapter new_dataAdapter2 = new SqlDataAdapter(ExelComm2);
                    new_dataAdapter2.Fill(ds2);
                }
                catch { }
                //MessageBox.Show(PfmName.ToString());
                //object[,] ReportDataArray2 = new object[ds2.Tables[0].Rows.Count, 1];

                ReportWorksheet.get_Range(CellName(4, 3), CellName(4, 3)).Value2 = PfmName;
                ReportWorksheet.get_Range(CellName(ds.Tables[0].Rows.Count + StartRow + 1, 3), CellName(ds.Tables[0].Rows.Count + StartRow + 1, 3)).Value2 = tb_Nach_Podr.Text;
                ExcelApp.GetType().InvokeMember("Run", System.Reflection.BindingFlags.Default | System.Reflection.BindingFlags.InvokeMethod, null, ExcelApp, new object[] { "Format" }); 

                // Отображение отчёта на экране
                ExcelApp.Visible = true;
                this.Close();
            }
        }

        private string CellName(int CellRow, int CellColumn)
        {
            return ((char)(64 + CellColumn) + CellRow.ToString());
        }

        
    }
}
