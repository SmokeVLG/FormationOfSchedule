using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Configuration;
//using System.Data.SqlClient;
//using System.Text;
using System.Windows.Forms;
using System.Data.Common;
using System.Data;
using System.Collections;
using System.ComponentModel;
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
using DevExpress.Utils;
using FormationOfSchedule.Properties;
using System.Data.Odbc;

namespace FormationOfSchedule
{
    class DataSaver
    {
        SqlConnection _connection;

        public DataSaver(string login, string password)
        {
            Connection(login, password);
        }

        public DataSaver(SqlConnection conn)
        {
            _connection = conn;
        }

        char separator_false;
        char separator_true;
        public void Connection(string login, string password)
        {
            if (_connection == null)
            {
                _connection = new SqlConnection(ConfigurationManager.ConnectionStrings["FormationOfSchedule.Properties.Settings.FormationOfSchedule"].ConnectionString);
                SqlConnectionStringBuilder scsb = new SqlConnectionStringBuilder();
                scsb.DataSource = _connection.DataSource;
                scsb.InitialCatalog = _connection.Database;
                scsb.UserID = login;
                scsb.Password = password;
                _connection = new SqlConnection(scsb.ConnectionString);
                //SqlConnectionStringBuilder scsb = new SqlConnectionStringBuilder();
                //scsb.DataSource = "budsccm";
               // scsb.InitialCatalog = "FormationOfSchedule";
                //scsb.UserID = login;
               // scsb.Password = password;
               // _connection = new SqlConnection(scsb.ConnectionString);
                login_ = login;

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
        }
        string login_;
        public string Save(System.Data.DataTable dt, int num_first_row)
        {
          //  Connection();
            string status = "";
            string statusALL = "";
            int k = 0;
            int kk = 0;
            int error_count = 0;
            var Warning_One = new List<string>(); 
            var Warning_Two = new List<string>();


            if (dt.Columns.Count != 22)
            {
                statusALL = "Ошибка: Загружен неправильный шаблон!";
            }
            else
            foreach ( DataRow row in dt.Rows)
            {

                if (row[17].ToString() == "")
                    row[17] = 0;

                k++; kk++;
                SqlCommand comm = new SqlCommand("usp_FS_information_to_DB", _connection);

                SqlParameter tParam = new SqlParameter("@Creditor", SqlDbType.NVarChar);
                tParam.Value = row[0].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@Contract", SqlDbType.NVarChar);
                tParam.Value = row[1].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@InDoc", SqlDbType.NVarChar);
                tParam.Value = row[2].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@ContractStatus", SqlDbType.NVarChar);
                tParam.Value = row[4].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@FinPosition", SqlDbType.NVarChar);
                tParam.Value = row[5].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@Name", SqlDbType.NVarChar);
                tParam.Value = row[6].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@Describe", SqlDbType.NVarChar);
                tParam.Value = row[7].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@PFMcode", SqlDbType.NVarChar);
                tParam.Value = row[8].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@PFMname", SqlDbType.NVarChar);
                tParam.Value = row[9].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@CuratorName", SqlDbType.NVarChar);
                tParam.Value = row[10].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@Executer", SqlDbType.NVarChar);
                tParam.Value = row[11].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@UrNumber", SqlDbType.NVarChar);
                tParam.Value = row[13].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@DateLog", SqlDbType.Date);
                tParam.Value = row[14];
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@DateStart", SqlDbType.Date);
                tParam.Value = row[15];
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@DateEnd", SqlDbType.Date);
                if(Convert.ToDateTime(row[16]).Year > 3000)
                    row[16] = Convert.ToDateTime(Convert.ToString(Convert.ToDateTime(row[16]).Day) + "." + Convert.ToString(Convert.ToDateTime(row[16]).Month) + ".3000");
                tParam.Value = row[16];
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@Currency", SqlDbType.NVarChar);
                tParam.Value = row[18].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@InvProjectCode", SqlDbType.NVarChar);
                tParam.Value = row[20].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@Summ", SqlDbType.Float);
                tParam.Value = row[17].ToString().Replace(separator_false, separator_true);
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@ProjectName", SqlDbType.NVarChar);
                tParam.Value = row[21].ToString();
                comm.Parameters.Add(tParam);
                comm.CommandType = System.Data.CommandType.StoredProcedure;

                tParam = new SqlParameter("@warning", SqlDbType.NVarChar);
                tParam.Direction = ParameterDirection.Output;
                tParam.Value = "";
                tParam.Size = 500;
                comm.Parameters.Add(tParam);
                
                try
                {
                    if (_connection.State != System.Data.ConnectionState.Open)
                        _connection.Open();
                    comm.ExecuteScalar();

                    status = comm.Parameters["@warning"].Value.ToString();
                  /*  if(status.ToString() != "")
                        Warning_One.Add(status);
*/


                    if (statusALL != status && status != "" && status.Substring(0, 6).ToString() != "Ошибка")
                    {
                        statusALL = statusALL + "\r\n" + 
                            //kk.ToString() + ".  " + 
                            status;
                        error_count++;
                    }
                    if (status != "" &&  status.Substring(0, 6).ToString() == "Ошибка")
                     {
                         statusALL = statusALL + "\r\n Excel#: " + (k + num_first_row - 1).ToString() + ".  " + status;
                         error_count++;
                         kk--;
                     }
               }
               catch 
                {
                   // MessageBox.Show((k + num_first_row - 1).ToString());
                    statusALL = "Ошибка: Возможно загружен неправильный шаблон! Внимательно осмотрите строку " + (k + num_first_row - 1).ToString() + " в Excel";
                }
               // catch (Exception ex)
              //  {
               //     MessageBox.Show(ex.Message);
              //  }

                Disconnection();
            }
           /*
            Warning_Two.Add(Warning_One[0]);
            foreach (var warn1 in Warning_One)
            {
                for (var warn2 = 0; warn2 < Warning_Two.Count; warn2++ )
                {
                    if (Warning_Two[warn2] != warn1)
                    {
                        Warning_Two.Add(warn1);
                    }
                }
            }
            

            foreach (var warn2 in Warning_Two)
            {
                statusALL = statusALL + warn2 + "\n";
            }
             */

            if (error_count > 6)
            {
                ErrorList(statusALL);
                return "";
            }
            else
                return statusALL;
        }


        private void ErrorList(string error)
        {
            string TemplateFileName = "ErrorList.txt";
           // string TemplatePath = Directory.GetCurrentDirectory() + "\\" + TemplateFileName;
            string TemplatePath = "C:\\WINDOWS\\Temp\\" + TemplateFileName;
            //System.IO.File.Create(TemplatePath);
            System.IO.File.WriteAllText(TemplatePath, error);
            //System.IO.File.AppendAllText(TemplatePath, error);

          //  StreamWriter sw = new StreamWriter(TemplatePath);
           // sw.WriteLine(error);
           // sw.Close(); 

            Process.Start(TemplatePath);

        }


        public string SaveLimits(System.Data.DataTable dt, int num_first_row)
        {
            int error = 0;
            string status = "";
            string statusALL = "";
            int k = 0;
            if (dt.Columns.Count != 6)
            {
                statusALL = "Ошибка: Загружен неправильный шаблон!";
            }
            else
            foreach ( DataRow row in dt.Rows)
            {
                if (row[5].ToString() == "")
                    row[5] = 0;

                k++;
                SqlCommand comm = new SqlCommand("usp_FS_SaveLimits", _connection);

                SqlParameter tParam = new SqlParameter("@Summ", SqlDbType.Float);
                tParam.Value = row[5].ToString().Replace(separator_false, separator_true);
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@PFMcode", SqlDbType.NVarChar);
                tParam.Value = row[0].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@FinPositionEPL", SqlDbType.NVarChar);
                tParam.Value = row[1].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@Year", SqlDbType.NVarChar);
                tParam.Value = row[3].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@Month", SqlDbType.NVarChar);
                tParam.Value = row[2].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@CurrencyRUB", SqlDbType.NVarChar);
                tParam.Value = row[4].ToString();
                comm.Parameters.Add(tParam);
               
                tParam = new SqlParameter("@warning", SqlDbType.NVarChar);
                tParam.Direction = ParameterDirection.Output;
                tParam.Value = "";
                tParam.Size = 500;
                comm.Parameters.Add(tParam);
               
                comm.CommandType = System.Data.CommandType.StoredProcedure;
                try
                {
                    if (_connection.State != System.Data.ConnectionState.Open)
                        _connection.Open();
                    comm.ExecuteScalar();

                    status = comm.Parameters["@warning"].Value.ToString();
                    if (statusALL != status && status != "")
                    {
                        error++;
                        statusALL = statusALL + "\r\n Excel#: " + (k + num_first_row - 1).ToString() + ".  " + status;
                    }
                  
                }
                catch 
                {
                    statusALL = "Ошибка: Возможно загружен неправильный шаблон!";
                }
                finally
                {
                    Disconnection();
                }
                
            }
            if (error > 1)
            {
                ErrorList(statusALL);
                return "";
            }
            else
                return statusALL;
        }


        public string SaveKSSStable(System.Data.DataTable dt, int num_first_row)
        {
            int error = 0;
            string status = "";
            string statusALL = "";
            int k = 0;
            if (dt.Columns.Count != 4)
            {
                statusALL = "Ошибка: Загружен неправильный шаблон!";
            }
            else
                foreach (DataRow row in dt.Rows)
                {

                    k++;
                    SqlCommand comm = new SqlCommand("usp_FS_SaveKSSStable", _connection);

                    SqlParameter tParam = new SqlParameter("@ContragentType", SqlDbType.NVarChar);
                    tParam.Value = row[0].ToString();
                    comm.Parameters.Add(tParam);

                    tParam = new SqlParameter("@ContragentName", SqlDbType.NVarChar);
                    tParam.Value = row[1].ToString();
                    comm.Parameters.Add(tParam);

                    tParam = new SqlParameter("@ContragentCode", SqlDbType.NVarChar);
                    tParam.Value = row[3].ToString();
                    comm.Parameters.Add(tParam);

                    tParam = new SqlParameter("@KSSScode", SqlDbType.NVarChar);
                    tParam.Value = row[2].ToString();
                    comm.Parameters.Add(tParam);

                    tParam = new SqlParameter("@warning", SqlDbType.NVarChar);
                    tParam.Direction = ParameterDirection.Output;
                    tParam.Value = "";
                    tParam.Size = 500;
                    comm.Parameters.Add(tParam);

                    comm.CommandType = System.Data.CommandType.StoredProcedure;
                    try
                    {
                        if (_connection.State != System.Data.ConnectionState.Open)
                            _connection.Open();
                        comm.ExecuteScalar();

                        status = comm.Parameters["@warning"].Value.ToString();
                        if (statusALL != status && status != "")
                        {
                            error++;
                            statusALL = statusALL + "\r\n Excel#: " + (k + num_first_row - 1).ToString() + ".  " + status;
                        }

                    }
                    catch
                    {
                        statusALL = "Ошибка: Возможно загружен неправильный шаблон!";
                    }
                    finally
                    {
                        Disconnection();
                    }

                }
            if (error > 1)
            {
                ErrorList(statusALL);
                return "";
            }
            else
                return statusALL;
        }

        public string SaveActualPayment(System.Data.DataTable dt, int num_first_row)
        {
            int error = 0;
            string status = "";
            string statusALL = "";
            int k = 0;
            if (dt.Columns.Count != 8)
            {
                statusALL = "Ошибка: Загружен неправильный шаблон!";
            }
            else

            foreach (DataRow row in dt.Rows)
            {
                if (row[6].ToString() == "")
                    row[6] = 0;
                if (row[7].ToString() == "")
                    row[7] = 0;

                k++;
                SqlCommand comm = new SqlCommand("usp_FB_SaveActualPayment", _connection);

                SqlParameter tParam = new SqlParameter("@Summ", SqlDbType.Float);
                tParam.Value = row[6].ToString().Replace(separator_false,separator_true);
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@SummRus", SqlDbType.Float);
                tParam.Value = row[7].ToString().Replace(separator_false, separator_true);
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@Contract", SqlDbType.NVarChar);
                tParam.Value = row[0].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@Creditor", SqlDbType.NVarChar);
                tParam.Value = row[1].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@FinPosition", SqlDbType.NVarChar);
                tParam.Value = row[2].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@PFMcode", SqlDbType.NVarChar);
                tParam.Value = row[3].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@DateLog", SqlDbType.DateTime);
                tParam.Value = row[4];
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@Currency", SqlDbType.NVarChar);
                tParam.Value = row[5].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@warning", SqlDbType.NVarChar);
                tParam.Direction = ParameterDirection.Output;
                tParam.Value = "";
                tParam.Size = 500;
                comm.Parameters.Add(tParam);

                comm.CommandType = System.Data.CommandType.StoredProcedure;
                try
               {
                    if (_connection.State != System.Data.ConnectionState.Open)
                        _connection.Open();
                    comm.ExecuteScalar();

                    status = comm.Parameters["@warning"].Value.ToString();
                    if (statusALL != status && status != "")
                    {
                        error++;
                        statusALL = statusALL + "\r\n Excel#: " + (k + num_first_row - 1).ToString() + ".  " + status;
                    }
                }
                catch 
                {
                    statusALL = "Ошибка: Возможно загружен неправильный шаблон!";
                }
                finally
                {
                    Disconnection();
                }

            }
            if (error > 1)
            {
                ErrorList(statusALL);
                return "";
            }
            else
                return statusALL;
        }

        public string CurrencyCursFill(System.Data.DataRow rrow)
        
        {
           // Connection();    

                SqlCommand comm = new SqlCommand("usp_update_curs", _connection);
                comm.CommandType = System.Data.CommandType.StoredProcedure;

               //MessageBox.Show("" + rrow["IdCurs"]);
            
                SqlParameter tParam = new SqlParameter("@DateCurs", SqlDbType.Date);
                tParam.Value = rrow["DateCurs"].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@Currency", SqlDbType.NVarChar);
                tParam.Value = rrow["Currency"].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@Rus", SqlDbType.Float);
                tParam.Value = rrow["Rus"].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@IdCurs", SqlDbType.Float);
                tParam.Value = rrow["IdCurs"].ToString();
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@res", SqlDbType.Char);
                tParam.Direction = ParameterDirection.Output;
                tParam.Value = "";
                tParam.Size = 50;
                comm.Parameters.Add(tParam);


                    if (_connection.State != System.Data.ConnectionState.Open)
                        _connection.Open();
                    comm.ExecuteNonQuery();

                    string status;
                    status = comm.Parameters["@res"].Value.ToString();

                    Disconnection();
                return status;
        }

        public string UpdateKSSStoContragent(string ContrType, string ContrName, string KSSS, string ContrCode, string id)
        { 
              //  Connection();

                SqlCommand comm = new SqlCommand("usp_FS_Update_KSSStoContragent", _connection);
                comm.CommandType = System.Data.CommandType.StoredProcedure;


                SqlParameter tParam = new SqlParameter("@ContragentType", SqlDbType.NVarChar);
                tParam.Value = ContrType;
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@id", SqlDbType.Int);
                tParam.Value = id;
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@ContragentName", SqlDbType.NVarChar);
                tParam.Value = ContrName;
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@KSSScode", SqlDbType.NVarChar);
                tParam.Value = KSSS;
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@ContragentCode", SqlDbType.NVarChar);
                tParam.Value = ContrCode;
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@res", SqlDbType.Char);
                tParam.Direction = ParameterDirection.Output;
                tParam.Value = "";
                tParam.Size = 50;
                comm.Parameters.Add(tParam);


                    if (_connection.State != System.Data.ConnectionState.Open)
                        _connection.Open();
                    comm.ExecuteNonQuery();

                    string status;
                    status = comm.Parameters["@res"].Value.ToString();

                    Disconnection();
                return status;
            
        }

        public string InsertKsssContragent(string ContrType, string ContrName, string KSSS, string ContrCode)
        {
           // Connection();

            SqlCommand comm = new SqlCommand("usp_FS_Add_to_KSSStoContragent", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@ContragentType", SqlDbType.NVarChar);
            tParam.Value = ContrType;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@ContragentName", SqlDbType.NVarChar);
            tParam.Value = ContrName;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@KSSScode", SqlDbType.NVarChar);
            tParam.Value = KSSS;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@ContragentCode", SqlDbType.NVarChar);
            tParam.Value = ContrCode;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@res", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 50;
            comm.Parameters.Add(tParam);


            if (_connection.State != System.Data.ConnectionState.Open)
                _connection.Open();
            comm.ExecuteNonQuery();

            string status;
            status = comm.Parameters["@res"].Value.ToString();

            Disconnection();
            return status;
        }

        public string DelKSSSContragent(string id)
        {
           // Connection();

            SqlCommand comm = new SqlCommand("usp_FS_Del_KSSStoContragent", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@id", SqlDbType.NVarChar);
            tParam.Value = id;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@res", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 50;
            comm.Parameters.Add(tParam);


            if (_connection.State != System.Data.ConnectionState.Open)
                _connection.Open();
            comm.ExecuteNonQuery();

            string status;
            status = comm.Parameters["@res"].Value.ToString();

            Disconnection();
            return status;
        }

        public string InsertPaymentPlan(string ContractCode, string PartnerCode, string PartnerType, string FinPosition, 
                                        string FinPositionEPL, string PFMcode, string DatePay, string curs,
                                        string PaySumm, string PaySummRus, string Comment)
        {
           // Connection();

            SqlCommand comm = new SqlCommand("usp_FS_Add_to_PaymentsPlan", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@ContractCode", SqlDbType.NVarChar);
            tParam.Value = ContractCode;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@PartnerCode", SqlDbType.NVarChar);
            tParam.Value = PartnerCode;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@PartnerType", SqlDbType.NVarChar);
            tParam.Value = PartnerType;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@FinPosition", SqlDbType.NVarChar);
            tParam.Value = FinPosition;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@FinPositionEPL", SqlDbType.NVarChar);
            tParam.Value = FinPositionEPL;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@PFMcode", SqlDbType.NVarChar);
            tParam.Value = PFMcode;
            comm.Parameters.Add(tParam);
            
            tParam = new SqlParameter("@DatePay", SqlDbType.DateTime);
            tParam.Value = DatePay;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@PayCurrency", SqlDbType.NVarChar);
            tParam.Value = curs;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@PaySumm", SqlDbType.Float);
            tParam.Value = PaySumm.Replace(separator_false, separator_true);
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@PaySummRus", SqlDbType.Float);
            tParam.Value = PaySummRus.Replace(" ", "");
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@Comment", SqlDbType.NVarChar);
            tParam.Value = Comment;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@res", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 200;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@resSumm", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 80;
            comm.Parameters.Add(tParam);
            try
            {
                if (_connection.State != System.Data.ConnectionState.Open)
                    _connection.Open();
                comm.ExecuteNonQuery();
            }
             catch (Exception ex)
              {
                 MessageBox.Show(ex.Message);
              }
            string status;
            string status1;
           // status = (string)comm.ExecuteScalar();
            status = comm.Parameters["@res"].Value.ToString();
            status1 = comm.Parameters["@resSumm"].Value.ToString();
            Disconnection();

            if (status1 == "")
                return status;
            else
                return status1;
        }

        public string UpdatePaymentPlan(string ContractCode, string PartnerCode, string PartnerType, string FinPosition,
                                        string FinPositionEPL, string PFMcode, string DatePay, string curs,
                                        string PaySumm, string PaySummRus, string id, string Comment, string admMet)
        {
            // Connection();

            SqlCommand comm = new SqlCommand("usp_FS_Update_PaymentPlans", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@ContractCode", SqlDbType.NVarChar);
            tParam.Value = ContractCode;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@PartnerCode", SqlDbType.NVarChar);
            tParam.Value = PartnerCode;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@PartnerType", SqlDbType.NVarChar);
            tParam.Value = PartnerType;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@FinPosition", SqlDbType.NVarChar);
            tParam.Value = FinPosition;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@FinPositionEPL", SqlDbType.NVarChar);
            tParam.Value = FinPositionEPL;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@PFMcode", SqlDbType.NVarChar);
            tParam.Value = PFMcode;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@DatePay", SqlDbType.DateTime);
            tParam.Value = DatePay;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@PayCurrency", SqlDbType.NVarChar);
            tParam.Value = curs;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@PaySumm", SqlDbType.Float);
            tParam.Value = PaySumm.Replace(separator_false, separator_true);
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@PaySummRus", SqlDbType.Float);
            tParam.Value = PaySummRus.Replace(separator_false, separator_true);
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@id", SqlDbType.Int);
            tParam.Value = id;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@Comment", SqlDbType.NVarChar);
            tParam.Value = Comment;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@admMet", SqlDbType.Int);
            tParam.Value = admMet;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@res", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 60;
            comm.Parameters.Add(tParam);

            try
            {
                if (_connection.State != System.Data.ConnectionState.Open)
                    _connection.Open();
                comm.ExecuteNonQuery();

            }
             catch (Exception ex)
             {
                 MessageBox.Show(ex.Message);
             }
                string status;
            // status = (string)comm.ExecuteScalar();
            status = comm.Parameters["@res"].Value.ToString();

            Disconnection();
            return status;
        }


        public string GetEPL(string FinPosition)
        {
           // Connection();

            SqlCommand comm = new SqlCommand("usp_FS_GET_FinPositionEPL", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@FinPosition", SqlDbType.NVarChar);
            tParam.Value = FinPosition;
            comm.Parameters.Add(tParam);


            if (_connection.State != System.Data.ConnectionState.Open)
                _connection.Open();
           // comm.ExecuteScalar();

            string EPL = (string)comm.ExecuteScalar();
            //EPL = comm.Parameters["@FinPosEPL"].Value.ToString();

            Disconnection();
            return EPL;
        }

        public string GetFinPositonName(string FinPosition)
        {
            // Connection();

            SqlCommand comm = new SqlCommand("usp_FS_GET_FinPositionName", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@FinPosition", SqlDbType.NVarChar);
            tParam.Value = FinPosition;
            comm.Parameters.Add(tParam);


            if (_connection.State != System.Data.ConnectionState.Open)
                _connection.Open();
            // comm.ExecuteScalar();

            string NameFP = (string)comm.ExecuteScalar();
            //EPL = comm.Parameters["@FinPosEPL"].Value.ToString();

            Disconnection();
            return NameFP;
        }

        public string GetStateEverydayLicvid(string FinPositionEPL)
        {

            SqlCommand comm = new SqlCommand("usp_FS_GET_StateEverydayLicvid", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@FinPositionEPL", SqlDbType.NVarChar);
            tParam.Value = FinPositionEPL;
            comm.Parameters.Add(tParam);


            if (_connection.State != System.Data.ConnectionState.Open)
                _connection.Open();
            string StateEverydayLicvid = (string)comm.ExecuteScalar();

            Disconnection();
            return StateEverydayLicvid;
        }

        public string GetContragentType(string ContragentCode)
        {
          //  Connection();

            SqlCommand comm = new SqlCommand("usp_FS_GET_ContragentType", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@ContragentCode", SqlDbType.NVarChar);
            tParam.Value = ContragentCode;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@res", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 50;
            comm.Parameters.Add(tParam);

            if (_connection.State != System.Data.ConnectionState.Open)
                _connection.Open();
            comm.ExecuteNonQuery();

            string status = comm.Parameters["@res"].Value.ToString();

            Disconnection();
            return status;
        }

        public string GetRUSsumm(string Summ, string currency, string data)
        {
            if (Summ == "")
                Summ = "0";
            SqlCommand comm = new SqlCommand("usp_FS_GET_SummRUB", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@Summ", SqlDbType.Float);
            tParam.Value = Summ.Replace(separator_false,separator_true);
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@Currency", SqlDbType.NVarChar);
            tParam.Value = currency;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@data", SqlDbType.Date);
            tParam.Value = data;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@res", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 50;
            comm.Parameters.Add(tParam);

            if (_connection.State != System.Data.ConnectionState.Open)
                _connection.Open();
            comm.ExecuteNonQuery();

            string status = comm.Parameters["@res"].Value.ToString().Replace(separator_false,separator_true);

            Disconnection();
            return status;
        }

        public string[] getPFM()
        {

                string[] LogInformation = new string[3];
                

                SqlCommand comm = new SqlCommand("usp_FS_LogInform", _connection);
                comm.CommandType = System.Data.CommandType.StoredProcedure;

                SqlParameter tParam = new SqlParameter("@login", SqlDbType.NVarChar);
                tParam.Value = login_;
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@PFMcode", SqlDbType.Char);
                tParam.Direction = ParameterDirection.Output;
                tParam.Value = "";
                tParam.Size = 20;
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@userGroup", SqlDbType.Char);
                tParam.Direction = ParameterDirection.Output;
                tParam.Value = "";
                tParam.Size = 20;
                comm.Parameters.Add(tParam);

                tParam = new SqlParameter("@blockMark", SqlDbType.Char);
                tParam.Direction = ParameterDirection.Output;
                tParam.Value = "";
                tParam.Size = 1;
                comm.Parameters.Add(tParam);

                if (_connection.State != System.Data.ConnectionState.Open)
                    _connection.Open();
                comm.ExecuteNonQuery();

                LogInformation[0] = comm.Parameters["@userGroup"].Value.ToString();
                LogInformation[1] = comm.Parameters["@PFMcode"].Value.ToString();
                LogInformation[2] = comm.Parameters["@blockMark"].Value.ToString();
                LogInformation_ = LogInformation;
                return LogInformation;
        }
        string[] LogInformation_ = new string[3];
        public void Select()
        {
            try
            {
                SqlCommand comm = new SqlCommand("select [PartnerCode] as col1 ,[ContractCode] as col2 ,[ContractStatus] as col3" +
                        " ,[FinPosition] as col4 ,[FinPositionEPL] as col19 ,[PartnersOrCreditorsName] as col5" +
                        " ,[DescriptionOfContracts] as col6 ,[PFMcode] as col7 ,[PFMname] as col8" +
                        " ,[CuratorName] as col9 ,[ExecuterName] as col10 ,[ContractNumber] as col11" +
                        " ,[DateLogContract] as col12 ,[DateStartContract] as col13 ,[DateEndContract] as col14" +
                        " ,[SummContract] as col15 ,[Currency] as col16 ,[InvestmentProjectCode] as col17" +
                        " ,[ProjectName] as col18 FROM [Contracts]", _connection);
                comm.ExecuteScalar();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public string AddLimits(string PFM, string finpos, string year, string month, string summ)
        {
            SqlCommand comm = new SqlCommand("usp_FS_Add_to_Limits", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@PFM", SqlDbType.NVarChar);
            tParam.Value = PFM;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@FinPosEPL", SqlDbType.NVarChar);
            tParam.Value = finpos;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@Year", SqlDbType.NVarChar);
            tParam.Value = year;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@Mounth", SqlDbType.NVarChar);
            tParam.Value = month;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@Summ", SqlDbType.Float);
            tParam.Value = summ.Replace(separator_false,separator_true);
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@res", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 50;
            comm.Parameters.Add(tParam);

            if (_connection.State != System.Data.ConnectionState.Open)
                _connection.Open();
            comm.ExecuteNonQuery();
            string res = "";
            res = comm.Parameters["@res"].Value.ToString();
            Disconnection();
            return res;
        }

        public string UpdateLimits(string PFM, string finpos, string year, string month, string summ, string id)
        {
            SqlCommand comm = new SqlCommand("usp_FS_Update_Limits", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@PFMcode", SqlDbType.NVarChar);
            tParam.Value = PFM;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@FinPositionEPL", SqlDbType.NVarChar);
            tParam.Value = finpos;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@Year", SqlDbType.NVarChar);
            tParam.Value = year;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@Month", SqlDbType.NVarChar);
            tParam.Value = month;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@Summ", SqlDbType.Float);
            tParam.Value = summ.Replace(separator_false, separator_true);
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@id", SqlDbType.Int);
            tParam.Value = id;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@res", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 50;
            comm.Parameters.Add(tParam);

            if (_connection.State != System.Data.ConnectionState.Open)
                _connection.Open();
            comm.ExecuteNonQuery();
            string res = "";
            res = comm.Parameters["@res"].Value.ToString();
            Disconnection();
            return res;
        }

        public string DeleteLimits(string idRow)
        {
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
            string res = "";
            res = comm.Parameters["@res"].Value.ToString();
            Disconnection();
            return res;
        }

        public string AddUsers(string login, string fio, string PFM, string PFMname, string group)
        {
            SqlCommand comm = new SqlCommand("usp_FS_Add_to_Users", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@Login", SqlDbType.NVarChar);
            tParam.Value =login;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@UsersName", SqlDbType.NVarChar);
            tParam.Value = fio;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@PFMcode", SqlDbType.NVarChar);
            tParam.Value = PFM;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@PFMname", SqlDbType.NVarChar);
            tParam.Value = PFMname;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@UsersGroup", SqlDbType.Int);
            tParam.Value = group;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@res", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 50;
            comm.Parameters.Add(tParam);

            if (_connection.State != System.Data.ConnectionState.Open)
                _connection.Open();
            comm.ExecuteNonQuery();
            string res = "";
            res = comm.Parameters["@res"].Value.ToString();
            Disconnection();
            return res;
        }

        public string UpdateUsers(string login, string fio, string PFM, string PFMname, string group, string id)
        {
            SqlCommand comm = new SqlCommand("usp_FS_Update_Users", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;


            SqlParameter tParam = new SqlParameter("@Login", SqlDbType.NVarChar);
            tParam.Value = login;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@id", SqlDbType.Int);
            tParam.Value = id;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@UsersName", SqlDbType.NVarChar);
            tParam.Value = fio;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@PFMcode", SqlDbType.NVarChar);
            tParam.Value = PFM;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@PFMname", SqlDbType.NVarChar);
            tParam.Value = PFMname;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@UsersGroup", SqlDbType.Int);
            tParam.Value = group;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@res", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 50;
            comm.Parameters.Add(tParam);

            if (_connection.State != System.Data.ConnectionState.Open)
                _connection.Open();
            comm.ExecuteNonQuery();
            string res = "";
            res = comm.Parameters["@res"].Value.ToString();
            Disconnection();
            return res;
        }

        public string DeleteUsers(string id)
        {
            SqlCommand comm = new SqlCommand("usp_FS_Del_in_Users", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@id", SqlDbType.Int);
            tParam.Value = id;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@res", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 50;
            comm.Parameters.Add(tParam);

            if (_connection.State != System.Data.ConnectionState.Open)
                _connection.Open();
            comm.ExecuteNonQuery();
            string res = "";
            res = comm.Parameters["@res"].Value.ToString();
            Disconnection();
            return res;
        }

        public string AddCurrency(string date, string curs, string currency_rub)
        {
            SqlCommand comm = new SqlCommand("usp_Add_Currency", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@DateCurs", SqlDbType.Date);
            tParam.Value = date;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@Currency", SqlDbType.NVarChar);
            tParam.Value = curs;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@Rus", SqlDbType.Float);
            tParam.Value = currency_rub;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@res", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 50;
            comm.Parameters.Add(tParam);
            string res = "";
            try
            {
                if (_connection.State != System.Data.ConnectionState.Open)
                    _connection.Open();
                comm.ExecuteNonQuery();
                
                res = comm.Parameters["@res"].Value.ToString();
            }
             catch (Exception ex)
              {
                 MessageBox.Show(ex.Message);
              }

            Disconnection();
            return res;
        }

        public string DeletePayments(string idRow)
        {
            SqlCommand comm = new SqlCommand("usp_FS_Del_in_PaymentPlans", _connection);
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
            string res = "";
            res = comm.Parameters["@res"].Value.ToString();
            Disconnection();
            return res;
        }

        public string InsertFinPosition(string FPcode, string StateEverydayLicvid)
        {
            // Connection();

            SqlCommand comm = new SqlCommand("usp_FS_Insert_FinancialPosition", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@FPcode", SqlDbType.NVarChar);
            tParam.Value = FPcode;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@StateEverydayLicvid", SqlDbType.NVarChar);
            tParam.Value = StateEverydayLicvid;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@res", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 50;
            comm.Parameters.Add(tParam);


            if (_connection.State != System.Data.ConnectionState.Open)
                _connection.Open();
            comm.ExecuteNonQuery();
            string status;
            // status = (string)comm.ExecuteScalar();
            status = comm.Parameters["@res"].Value.ToString();

            Disconnection();
            return status;
        }

        public string UpdateFinPosition(string FPcodeOLD, string FPcode, string StateEverydayLicvid)
        {
            // Connection();

            SqlCommand comm = new SqlCommand("usp_FS_Update_FinancialPosition", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@FPcode", SqlDbType.NVarChar);
            tParam.Value = FPcode;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@FPcodeOLD", SqlDbType.NVarChar);
            tParam.Value = FPcodeOLD;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@StateEverydayLicvid", SqlDbType.NVarChar);
            tParam.Value = StateEverydayLicvid;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@res", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 50;
            comm.Parameters.Add(tParam);


            if (_connection.State != System.Data.ConnectionState.Open)
                _connection.Open();
            comm.ExecuteNonQuery();
            string status;
            // status = (string)comm.ExecuteScalar();
            status = comm.Parameters["@res"].Value.ToString();

            Disconnection();
            return status;
        }

        public string UpdateFinPosition_Stavrolen(string FPcodeStavrolenOLD, string FPcodeStavrolen, string FPStavrolenName)
        {
            // Connection();

            SqlCommand comm = new SqlCommand("usp_FS_Update_FinancialPosition_Stavrolen", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@FPcodeStavrolen", SqlDbType.NVarChar);
            tParam.Value = FPcodeStavrolen;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@FPcodeStavrolenOLD", SqlDbType.NVarChar);
            tParam.Value = FPcodeStavrolenOLD;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@FPStavrolenName", SqlDbType.NVarChar);
            tParam.Value = FPStavrolenName;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@res", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 50;
            comm.Parameters.Add(tParam);


            if (_connection.State != System.Data.ConnectionState.Open)
                _connection.Open();
            comm.ExecuteNonQuery();
            string status;
            // status = (string)comm.ExecuteScalar();
            status = comm.Parameters["@res"].Value.ToString();

            Disconnection();
            return status;
        }

        public string DeleteFinPosition( string FPcode, string StateEverydayLicvid)
        {
            // Connection();

            SqlCommand comm = new SqlCommand("usp_FS_Del_FinancialPosition", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@FPcode", SqlDbType.NVarChar);
            tParam.Value = FPcode;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@StateEverydayLicvid", SqlDbType.NVarChar);
            tParam.Value = StateEverydayLicvid;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@res", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 50;
            comm.Parameters.Add(tParam);


            if (_connection.State != System.Data.ConnectionState.Open)
                _connection.Open();
            comm.ExecuteNonQuery();
            string status;
            // status = (string)comm.ExecuteScalar();
            status = comm.Parameters["@res"].Value.ToString();

            Disconnection();
            return status;
        }

        public string DeleteFinPosition_Stavrolen(string FPcodeStavrolen, string FPStavrolenName)
        {
            // Connection();

            SqlCommand comm = new SqlCommand("usp_FS_Del_FinancialPosition_Stavrolen", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@FPcodeStavrolen", SqlDbType.NVarChar);
            tParam.Value = FPcodeStavrolen;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@FPStavrolenName", SqlDbType.NVarChar);
            tParam.Value = FPStavrolenName;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@res", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 50;
            comm.Parameters.Add(tParam);


            if (_connection.State != System.Data.ConnectionState.Open)
                _connection.Open();
            comm.ExecuteNonQuery();
            string status;
            // status = (string)comm.ExecuteScalar();
            status = comm.Parameters["@res"].Value.ToString();

            Disconnection();
            return status;
        }

        public string InsertFinPosition_Stavrolen(string FPcode, string FPcodeStavrolen, string FPStavrolenName)
        {
            // Connection();

            SqlCommand comm = new SqlCommand("usp_FS_Insert_FinancialPositionStavrolen", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@FPcode", SqlDbType.NVarChar);
            tParam.Value = FPcode;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@FPcodeStavrolen", SqlDbType.NVarChar);
            tParam.Value = FPcodeStavrolen;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@FPStavrolenName", SqlDbType.NVarChar);
            tParam.Value = FPStavrolenName;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@res", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 50;
            comm.Parameters.Add(tParam);


            if (_connection.State != System.Data.ConnectionState.Open)
                _connection.Open();
            comm.ExecuteNonQuery();
            string status;
            // status = (string)comm.ExecuteScalar();
            status = comm.Parameters["@res"].Value.ToString();

            Disconnection();
            return status;
        }

        public string InsertPaymentsPlanSum(DateTime date)
        {
            string result = "";

            SqlCommand command = new SqlCommand("usp_FS_InsertPaymentPlanSum", _connection);
            command.CommandType = CommandType.StoredProcedure;

            command.Parameters.Add("@date", SqlDbType.Date).Direction = ParameterDirection.Input;
            command.Parameters["@date"].Value = date;

            command.Parameters.Add("@message", SqlDbType.VarChar, 300).Direction = ParameterDirection.Output;

            try
            {
                if (_connection.State != System.Data.ConnectionState.Open)
                    _connection.Open();
                command.ExecuteNonQuery();

                result = string.Join("\n", command.Parameters["@message"].Value.ToString().Split('|'));
            }
            catch (SqlException ex)
            {
                result = ex.Message;
            }

            Disconnection();

            return result;
        }

        public string SortFinPos(string id_tec, string id_)
        {
            SqlCommand comm = new SqlCommand("usp_FS_MovePosition", _connection);
            comm.CommandType = System.Data.CommandType.StoredProcedure;

            SqlParameter tParam = new SqlParameter("@id_tec", SqlDbType.Int);
            tParam.Value = id_tec;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@id_", SqlDbType.Int);
            tParam.Value = id_;
            comm.Parameters.Add(tParam);

            tParam = new SqlParameter("@res", SqlDbType.Char);
            tParam.Direction = ParameterDirection.Output;
            tParam.Value = "";
            tParam.Size = 50;
            comm.Parameters.Add(tParam);


            if (_connection.State != System.Data.ConnectionState.Open)
                _connection.Open();
            comm.ExecuteNonQuery();
            string status;
            // status = (string)comm.ExecuteScalar();
            status = comm.Parameters["@res"].Value.ToString();

            Disconnection();
            return status;
        }

        public string DeleteContracts(string idRow)
        {
            SqlCommand comm = new SqlCommand("usp_FS_Del_in_Contracts", _connection);
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
            string res = "";
            res = comm.Parameters["@res"].Value.ToString();
            Disconnection();
            return res;
        }

        public string DeleteActualPayments(string idRow)
        {
            SqlCommand comm = new SqlCommand("usp_FS_Del_in_ActualPayments", _connection);
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
            string res = "";
            res = comm.Parameters["@res"].Value.ToString();
            Disconnection();
            return res;
        }

        public void SavePaymentsPlan(System.Data.DataTable dt)
        {
                foreach (DataRow row in dt.Rows)
                {

                    if (row[9].ToString() == "")
                        row[9] = 0;

                    SqlCommand comm = new SqlCommand("usp_FS_Add_to_PaymentsPlan", _connection);
                    comm.CommandType = CommandType.StoredProcedure;

                    SqlParameter tParam = new SqlParameter("@ContractCode", SqlDbType.NVarChar,20);
                    tParam.Value = row[0].ToString();
                    comm.Parameters.Add(tParam);

                   // comm.Parameters.Add(new SqlParameter("@ContractCode", SqlDbType.VarChar, 20));
                   // comm.Parameters["@ContractCode"].Value = row[0].ToString();

                     tParam = new SqlParameter("@PartnerCode", SqlDbType.NVarChar);
                    tParam.Value = row[1].ToString();
                    comm.Parameters.Add(tParam);

                    tParam = new SqlParameter("@PartnerType", SqlDbType.NVarChar);
                    tParam.Value = row[2].ToString();
                    comm.Parameters.Add(tParam);

                    tParam = new SqlParameter("@FinPosition", SqlDbType.NVarChar);
                    tParam.Value = row[3].ToString();
                    comm.Parameters.Add(tParam);

                    tParam = new SqlParameter("@FinPositionEPL", SqlDbType.NVarChar);
                    tParam.Value = row[4].ToString();
                    comm.Parameters.Add(tParam);

                    tParam = new SqlParameter("@PFMcode", SqlDbType.NVarChar);
                    tParam.Value = row[5].ToString();
                    comm.Parameters.Add(tParam);

                    tParam = new SqlParameter("@DatePay", SqlDbType.Date);
                    tParam.Value = row[6];
                    comm.Parameters.Add(tParam);

                    tParam = new SqlParameter("@PayCurrency", SqlDbType.NVarChar);
                    tParam.Value = row[7].ToString();
                    comm.Parameters.Add(tParam);

                    tParam = new SqlParameter("@PaySumm", SqlDbType.Float);
                    tParam.Value = row[8].ToString().Replace(separator_false, separator_true);
                    comm.Parameters.Add(tParam);

                    tParam = new SqlParameter("@PaySummRus", SqlDbType.Float);
                    tParam.Value = row[9].ToString().Replace(separator_false, separator_true);
                    comm.Parameters.Add(tParam);


                    try
                    {
                        if (_connection.State != System.Data.ConnectionState.Open)
                            _connection.Open();
                        comm.ExecuteNonQuery();

                    }
                    catch
                    {

                    }

                    Disconnection();
                }
        }


        public void Disconnection()
        {
            _connection.Close();
        }


    }
}
