using ClosedXML.Excel;
using Newtonsoft.Json;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ConsoleApp3
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {

                if (DateTime.Now.Hour == Int32.Parse("11"))
                {
                    try
                    {
                        string EntranceId = "";
                        string ExitId = "";
                        Params(out EntranceId, out ExitId);
                        string constr = "DATA SOURCE=127.0.0.1:1521/orcl.local.com; PASSWORD=123;PERSIST SECURITY INFO=True;USER ID=SCU";//conection sstring
                        string id = "";
                        OracleConnection con = new OracleConnection();
                        con.ConnectionString = constr;
                        string select = Query();
                        select = select.Replace("ExitId", ExitId);
                        select = select.Replace("EnteranceId", EntranceId);
                        OracleCommand cmd = new OracleCommand("FLEX_INSERT_CMD", con);
                        cmd.Parameters.Add("ins_cmd", OracleDbType.Varchar2).Value = select;
                        cmd.CommandType = CommandType.StoredProcedure;
                        OracleParameter outparam = new OracleParameter();
                        outparam = cmd.Parameters.Add("CURS", OracleDbType.RefCursor, ParameterDirection.Output);
                        try
                        {
                            con.Open();
                            OracleDataReader rd = cmd.ExecuteReader();
                            while (rd.Read())
                            {
                                id = rd.GetValue(0).ToString();
                            }
                            Console.WriteLine("добавлен запрос");
                            con.Close();
                        }
                        catch (Exception r)
                        {
                            Console.WriteLine(r);
                            con.Close();
                        }
                        string resp = GetRes(id);
                        List<Employee> employees = JsonConvert.DeserializeObject<List<Employee>>(resp);
                        bool flag = false;


                        DateTime date1 = new DateTime();
                        date1 = DateTime.Now;

                        using (var workbook = new XLWorkbook())
                        {
                            IXLWorksheet worksheet = workbook.Worksheets.Add("Report");
                           
                            //worksheet.Cell(1, 1).Value = "Приход-уход сотрудников за: " + employees[0].DATE_DAY.ToString();
                            worksheet.Cell(2, 1).Value = "ФИО";
                            worksheet.Cell(2, 2).Value = "Наименование управления";
                            worksheet.Cell(2, 3).Value = "Время прихода";
                            worksheet.Cell(2, 4).Value = "Время ухода";
                            worksheet.Cell(2, 5).Value = "Время опоздания";
                            worksheet.Cell(2, 8).Value = "Уход на обед";
                            worksheet.Cell(2, 9).Value = "Приход с обеда";
                            //worksheet.Cell(2, 10).Value = employees[0].DATE_DAY.ToString();
                            int j = 3;
                            //if (employees[0].GROUP_NAME == "ВК Лексус" || employees[0].GROUP_NAME == "Внешние аудиторы"
                            //    || employees[0].GROUP_NAME == "Сберкассы" || employees[0].GROUP_NAME == "СК Берекет Гранд"
                            //    || employees[0].GROUP_NAME == "СК Табылга" || employees[0].GROUP_NAME == "Токо Тукан"
                            //    || employees[0].GROUP_NAME == "ХБК-Азия" || employees[0].GROUP_NAME == "ХБК-Кенч"
                            //    || employees[0].GROUP_NAME == "ХБК-ЮГ" || employees[0].GROUP_NAME == "яУволенные") { }
                            //else
                            //{
                            if (employees[1].id != employees[0].id)
                            {
                                if (String.IsNullOrEmpty(employees[0].EVENT_TYPE_ID))
                                {
                                    worksheet.Cell(3, 1).Value = employees[0].FIRST_NAME + " " + employees[0].MIDDLE_NAME + " " + employees[0].LAST_NAME;
                                    worksheet.Cell(3, 2).Value = employees[0].GROUP_NAME;
                                    worksheet.Cell(3, 3).Value = "Отсутствовал";
                                    worksheet.Cell(3, 4).Value = "Отсутствовал";

                                    worksheet.Cell(j, 8).Value = "Отсутствовал";
                                    worksheet.Cell(j, 9).Value = "Отсутствовал";

                                    worksheet.Cell(2, 10).Value = employees[0].DATE_DAY.ToString();
                                    j++;
                                }
                                else if (employees[0].EVENT_TYPE_ID == ExitId)
                                {
                                    worksheet.Cell(3, 1).Value = employees[0].FIRST_NAME + " " + employees[0].MIDDLE_NAME + " " + employees[0].LAST_NAME;
                                    worksheet.Cell(3, 2).Value = employees[0].GROUP_NAME;
                                    worksheet.Cell(3, 3).Value = "-";
                                    worksheet.Cell(3, 4).Value = employees[0].DATE_TIME;

                                    worksheet.Cell(j, 8).Value = "не уход";
                                    worksheet.Cell(j, 9).Value = "не приход";

                                    worksheet.Cell(j, 10).Value = employees[0].DATE_DAY.ToString();
                                    j++;
                                }
                                else if (employees[0].EVENT_TYPE_ID == EntranceId)
                                {
                                    worksheet.Cell(3, 1).Value = employees[0].FIRST_NAME + " " + employees[0].MIDDLE_NAME + " " + employees[0].LAST_NAME;
                                    worksheet.Cell(3, 2).Value = employees[0].GROUP_NAME;
                                    worksheet.Cell(3, 3).Value = employees[0].DATE_TIME;
                                    worksheet.Cell(3, 4).Value = "-";

                                    worksheet.Cell(j, 8).Value = "не уход";
                                    worksheet.Cell(j, 9).Value = "не приход";

                                    worksheet.Cell(j, 10).Value = employees[0].DATE_DAY.ToString();
                                    j++;

                                }
                            }
                            else if (employees[0].EVENT_TYPE_ID == EntranceId)
                            {
                                worksheet.Cell(3, 1).Value = employees[0].FIRST_NAME + " " + employees[0].MIDDLE_NAME + " " + employees[0].LAST_NAME;
                                worksheet.Cell(3, 2).Value = employees[0].GROUP_NAME;
                                worksheet.Cell(3, 3).Value = employees[0].DATE_TIME;
                                worksheet.Cell(3, 4).Value = "-";

                                worksheet.Cell(j, 10).Value = employees[0].DATE_DAY.ToString();

                            }
                            int i;
                            for (i = 1; i < employees.Count - 1; i++)
                            {

                                if (employees[i].id != employees[i - 1].id && employees[i].id != employees[i + 1].id)
                                {
                                    if (String.IsNullOrEmpty(employees[i].EVENT_TYPE_ID))
                                    {
                                        worksheet.Cell(j, 1).Value = employees[i].FIRST_NAME + " " + employees[i].MIDDLE_NAME + " " + employees[i].LAST_NAME;
                                        worksheet.Cell(j, 2).Value = employees[i].GROUP_NAME;
                                        worksheet.Cell(j, 3).Value = "Отсутствовал";
                                        worksheet.Cell(j, 4).Value = "Отсутствовал";

                                        worksheet.Cell(j, 8).Value = "Отсутствовал";
                                        worksheet.Cell(j, 9).Value = "Отсутствовал";

                                        worksheet.Cell(j, 10).Value = employees[0].DATE_DAY.ToString();

                                        j++;
                                    }
                                    else if (employees[i].EVENT_TYPE_ID == ExitId)
                                    {
                                        worksheet.Cell(j, 1).Value = employees[i].FIRST_NAME + " " + employees[i].MIDDLE_NAME + " " + employees[i].LAST_NAME;
                                        worksheet.Cell(j, 2).Value = employees[i].GROUP_NAME;
                                        worksheet.Cell(j, 3).Value = "-";
                                        worksheet.Cell(j, 4).Value = employees[i].DATE_TIME;

                                        worksheet.Cell(j, 8).Value = "не уход";
                                        worksheet.Cell(j, 9).Value = "не приход";



                                        j++;
                                    }
                                    else if (employees[i].EVENT_TYPE_ID == EntranceId)
                                    {
                                        worksheet.Cell(j, 1).Value = employees[i].FIRST_NAME + " " + employees[i].MIDDLE_NAME + " " + employees[i].LAST_NAME;
                                        worksheet.Cell(j, 2).Value = employees[i].GROUP_NAME;
                                        worksheet.Cell(j, 3).Value = employees[i].DATE_TIME;
                                        worksheet.Cell(j, 4).Value = "-";

                                        worksheet.Cell(j, 8).Value = "не уход";
                                        worksheet.Cell(j, 9).Value = "не приход";
                                        j++;
                                    }
                                    flag = true;
                                }
                                else if (employees[i].id != employees[i - 1].id)
                                {
                                    if (employees[i].EVENT_TYPE_ID == EntranceId)
                                    {
                                        worksheet.Cell(j, 1).Value = employees[i].FIRST_NAME + " " + employees[i].MIDDLE_NAME + " " + employees[i].LAST_NAME;
                                        worksheet.Cell(j, 2).Value = employees[i].GROUP_NAME;
                                        worksheet.Cell(j, 3).Value = employees[i].DATE_TIME;
                                        worksheet.Cell(j, 4).Value = "-";

                                        worksheet.Cell(j, 8).Value = "не уход";
                                        worksheet.Cell(j, 9).Value = "не приход";
                                        flag = false;
                                        //---------------------------------------------------


                                    }
                                    else if (employees[i].EVENT_TYPE_ID == ExitId)
                                    {
                                        worksheet.Cell(j, 1).Value = employees[i].FIRST_NAME + " " + employees[i].MIDDLE_NAME + " " + employees[i].LAST_NAME;
                                        worksheet.Cell(j, 2).Value = employees[i].GROUP_NAME;
                                        worksheet.Cell(j, 3).Value = "-";
                                        worksheet.Cell(j, 4).Value = employees[i].DATE_TIME;

                                        worksheet.Cell(j, 8).Value = "не уход";
                                        worksheet.Cell(j, 9).Value = "не приход";
                                        flag = false;
                                        //------------------------------------------








                                    }
                                    // обед

                                }
                                // может и не быть так как один сотрудник
                                else if (employees[i].id != employees[i + 1].id)
                                {
                                    if (employees[i].EVENT_TYPE_ID == ExitId)
                                    {
                                        worksheet.Cell(j, 1).Value = employees[i].FIRST_NAME + " " + employees[i].MIDDLE_NAME + " " + employees[i].LAST_NAME;
                                        worksheet.Cell(j, 2).Value = employees[i].GROUP_NAME;
                                        worksheet.Cell(j, 4).Value = employees[i].DATE_TIME;
                                        if (worksheet.Cell(j, 3).Value == worksheet.Cell(j, 10).Value)
                                            worksheet.Cell(j, 3).Value = "-";
                                        flag = true;
                                    }
                                    j++;


                                }

                                else if (flag && employees[i].EVENT_TYPE_ID == EntranceId)
                                {
                                    worksheet.Cell(j, 1).Value = employees[i].FIRST_NAME + " " + employees[i].MIDDLE_NAME + " " + employees[i].LAST_NAME;
                                    worksheet.Cell(j, 2).Value = employees[i].GROUP_NAME;
                                    worksheet.Cell(j, 3).Value = employees[i].DATE_TIME;
                                    flag = false;
                                }
                                else
                                {
                                    // обед
                                    GetDiner(worksheet, employees, i, j);
                                }
                                
                                string eppsTime = "";
                                try
                                {
                                    //string tsss = worksheet.Cell(j, 3).Value.ToString();

                                    //string tss = Convert.ToDateTime(worksheet.Cell(j, 3).Value.ToString()).ToString();//.TimeOfDay.ToString();

                                    if (Convert.ToDateTime(worksheet.Cell(j, 3).Value.ToString()).TimeOfDay >= Convert.ToDateTime("09:00:00").TimeOfDay)
                                    {
                                        TimeSpan ts = (Convert.ToDateTime(worksheet.Cell(j, 3).Value.ToString()).TimeOfDay - Convert.ToDateTime("09:00:00").TimeOfDay);
                                        eppsTime = $"{ts.Hours.ToString().Replace("-", "").PadLeft(2, '0')}:{ts.Minutes.ToString().Replace("-", "").PadLeft(2, '0')}:{ts.Seconds.ToString().Replace("-", "").PadLeft(2, '0')}";//Minutes.ToString();
                                        eppsTime = eppsTime;
                                    }
                                    else{
                                        eppsTime = "";
                                    }
                                }
                                catch (Exception ex)
                                {

                                }
                                //string per = (Convert.ToDateTime(worksheet.Cell(j, 3).Value).TimeOfDay - Convert.ToDateTime("09:00:00").TimeOfDay).ToString();
                                if (eppsTime != "")
                                {
                                    worksheet.Cell(j, 5).Value = eppsTime;
                                }
                                else
                                {
                                    worksheet.Cell(j, 5).Value = "-";
                                }

                            }
                            if (employees[i].id != employees[i - 1].id)
                            {
                                if (String.IsNullOrEmpty(employees[i].EVENT_TYPE_ID))
                                {
                                    worksheet.Cell(3, 1).Value = employees[i].FIRST_NAME + " " + employees[i].MIDDLE_NAME + " " + employees[i].LAST_NAME;
                                    worksheet.Cell(3, 2).Value = employees[i].GROUP_NAME;
                                    worksheet.Cell(3, 3).Value = "Отсутствовал";
                                    worksheet.Cell(3, 4).Value = "Отсутствовал";
                                    j++;
                                }
                                else if (employees[i].EVENT_TYPE_ID == ExitId)
                                {

                                    worksheet.Cell(3, 1).Value = employees[i].FIRST_NAME + " " + employees[i].MIDDLE_NAME + " " + employees[i].LAST_NAME;
                                    worksheet.Cell(3, 2).Value = employees[i].GROUP_NAME;
                                    worksheet.Cell(3, 3).Value = "-";
                                    worksheet.Cell(3, 4).Value = employees[i].DATE_TIME;
                                    j++;
                                }
                                else if (employees[i].EVENT_TYPE_ID == EntranceId)
                                {
                                    worksheet.Cell(3, 1).Value = employees[i].FIRST_NAME + " " + employees[i].MIDDLE_NAME + " " + employees[i].LAST_NAME;
                                    worksheet.Cell(3, 2).Value = employees[i].GROUP_NAME;
                                    worksheet.Cell(3, 3).Value = employees[i].DATE_TIME;
                                    worksheet.Cell(3, 4).Value = "-";
                                    j++;
                                }
                            }
                            else if (employees[i].EVENT_TYPE_ID == ExitId)
                            {
                                worksheet.Cell(j, 1).Value = employees[i].FIRST_NAME + " " + employees[i].MIDDLE_NAME + " " + employees[i].LAST_NAME;
                                worksheet.Cell(j, 2).Value = employees[i].GROUP_NAME;
                                worksheet.Cell(j, 4).Value = employees[i].DATE_TIME;
                                string eppsTime = "";
                                try
                                {
                                    double totalMin = (Convert.ToDateTime(worksheet.Cell(j, 3).Value.ToString()).TimeOfDay - Convert.ToDateTime("09:00:00").TimeOfDay).TotalMinutes;
                                    if (totalMin > 0)
                                        eppsTime = (Convert.ToDateTime(worksheet.Cell(j, 3).Value.ToString()).TimeOfDay - Convert.ToDateTime("09:00:00").TimeOfDay).ToString();
                                }
                                catch (Exception ex)
                                {

                                }
                               
                                if (eppsTime != "")
                                {
                                    worksheet.Cell(j, 5).Value = eppsTime;
                                }
                                else
                                {
                                    worksheet.Cell(j, 5).Value = "-";
                                }
                            }

                            
                            string InsertString = "INSERT ALL ";
                            for (j = 3; j <= worksheet.RowCount(); j++)
                            {
                                if (worksheet.Cell(j, 1).Value.ToString() != "")
                                {
                                    InsertString += "INTO TIMETB (DATES, NAMES, GROUPNAME, TIMECOME, LATE, LUNCHLEAVE, LUNCHCOME, TIMELEAVE) VALUES";
                                    if (worksheet.Cell(j, 10).Value.ToString() != "")
                                    {
                                        InsertString += 
                                            "('" + employees[0].DATE_DAY.ToString() +
                                            "', '" + worksheet.Cell(j, 1).Value.ToString() +
                                            "', '" + worksheet.Cell(j, 2).Value.ToString();
                                    }
                                    else
                                    {
                                        InsertString += 
                                            "('" + employees[0].DATE_DAY.ToString() +
                                            "', '" + worksheet.Cell(j, 1).Value.ToString() +
                                            "', '" + worksheet.Cell(j, 2).Value.ToString();
                                    }

                                    if (worksheet.Cell(j, 3).Value.ToString() != "Отсутствовал" || worksheet.Cell(j, 3).Value.ToString() != "-")
                                    {
                                        InsertString += "', '" + worksheet.Cell(j, 3).Value.ToString();
                                    }
                                    else
                                    {
                                        InsertString += "', '" + null ;
                                    }

                                    if (worksheet.Cell(j, 5).Value.ToString() != "-")
                                    {
                                        InsertString += "', '" + worksheet.Cell(j, 5).Value.ToString();
                                    }
                                    else
                                    {

                                        InsertString += "', '" + null;
                                    }

                                    if (worksheet.Cell(j, 8).Value.ToString() != "не уход" || worksheet.Cell(j, 8).Value.ToString() != "Отсутствовал")
                                    {
                                        InsertString += "', '" + worksheet.Cell(j, 8).Value.ToString();

                                    }
                                    else
                                    {
                                        InsertString += "', '" + null;
                                    }
                                    if (worksheet.Cell(j, 9).Value.ToString() != "не приход" || worksheet.Cell(j, 9).Value.ToString() != "Отсутствовал")
                                    {
                                        InsertString += "', '" + worksheet.Cell(j, 9).Value.ToString();

                                    }
                                    else
                                    {
                                        InsertString += "', '" + "null" + " ";
                                    }

                                    if (worksheet.Cell(j, 4).Value.ToString() != "Отсутствовал" || worksheet.Cell(j, 4).Value.ToString() != "-")
                                    {
                                        InsertString += "', '" + worksheet.Cell(j, 4).Value.ToString() + "') ";
                                    }
                                    else
                                    {
                                        InsertString += "', '" + "null" +"') ";
                                    }
                                }
                            }

                            InsertString += "SELECT 1 FROM DUAL;";

                            string conne = "DATA SOURCE=127.0.0.1:1521/orcl.local.com;PASSWORD=123;PERSIST SECURITY INFO=True;USER ID=SCU";//query con
                            OracleConnection conetc = new OracleConnection();
                            conetc.ConnectionString = conne;
                            OracleCommand cmds = new OracleCommand(InsertString, conetc);
                            cmds.CommandType = CommandType.Text;
                            try
                            {
                                conetc.Open();
                                cmds.ExecuteNonQuery();
                                conetc.Close();
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e);
                                conetc.Close();
                                throw (e);
                            }


                            using (var stream = new MemoryStream()) 
                            {
                                string filename = $@"Посещаемость_за {date1.AddDays(-1).ToString("dd.MM.yyyy")}.xlsx";

                                FileStream file = new FileStream(filename, FileMode.Create, System.IO.FileAccess.Write);
                                workbook.SaveAs(stream);


                                stream.WriteTo(file);
                                file.Close();

                               
                            }
                        }
                        
                    }
                    catch (Exception r)
                    {
                        Console.WriteLine(r);
                    }
                }
                Thread.Sleep(3600000);
            }
        }
        public static string Query()
        {
            string constr = "DATA SOURCE=127.0.0.1:1521/orcl.local.com;PASSWORD=123;PERSIST SECURITY INFO=True;USER ID=SCU";//query con
            OracleConnection con = new OracleConnection();
            con.ConnectionString = constr;
            OracleCommand cmd = new OracleCommand("Select QUERY from QUERIES where id = 21 ", con);
            cmd.CommandType = CommandType.Text;
            string query = "";
            try
            {
                con.Open();
                OracleDataReader rd = cmd.ExecuteReader();
                while (rd.Read())
                {
                    query = rd.GetValue(0).ToString();
                }
                Console.WriteLine("добавлен файл");
                con.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                con.Close();
                throw (e);
            }
            return query;
        }

        public static void Params(out string EnteranceId, out string ExitId)
        {
            string constr = "DATA SOURCE=127.0.0.1:1521/orcl.local.com;PASSWORD=123;PERSIST SECURITY INFO=True;USER ID=SCU";//queri con
            OracleConnection con = new OracleConnection();
            con.ConnectionString = constr;
            OracleCommand cmd = new OracleCommand("Select PARAM_NAME, PARAM_VALUE from PARAMS where QUERY_ID = 21   ", con);
            cmd.CommandType = CommandType.Text;
            EnteranceId = "";
            ExitId = "";
            try
            {
                con.Open();
                OracleDataReader rd = cmd.ExecuteReader();
                while (rd.Read())
                {
                    if (rd.GetValue(0).ToString() == "ExitId")
                    {
                        ExitId = rd.GetValue(1).ToString();
                    }
                    else if (rd.GetValue(0).ToString() == "EnteranceId")
                    {
                        EnteranceId = rd.GetValue(1).ToString();
                    }
                }
                Console.WriteLine("добавлен файл");
                con.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                con.Close();
                throw (e);
            }
        }


        public static void GetDiner (IXLWorksheet worksheet, List<Employee> employees, int i, int j)
        {
            var dateEnter = new TimeSpan();
            var hour1 = new TimeSpan(13, 20, 00);
            var hour2 = new TimeSpan(14, 30, 00);

            if (employees[i].DATE_TIME != null)
            {
                dateEnter = TimeSpan.Parse(employees[i].DATE_TIME);
                if (employees[i].EVENT_TYPE_ID == "101")
                {
                    if (dateEnter > hour1 && dateEnter < hour2)
                    {
                        //приход с обеда 
                         worksheet.Cell(j, 9).Value = employees[i].DATE_TIME;
                    }
                    else
                    {
                        //приход с обеда пустой
                         worksheet.Cell(j, 9).Value = "-";
                    }
                }
            }
            else
            {
                //дата пустая 
                 worksheet.Cell(j, 9).Value = "--";

            }


            var dateExit = new TimeSpan();
            var hourex1 = new TimeSpan(12, 50, 00);
            var hourex2 = new TimeSpan(13, 30, 00);

            if (employees[i].DATE_TIME != null)
            {
                dateExit = TimeSpan.Parse(employees[i].DATE_TIME);


                if (employees[i].EVENT_TYPE_ID == "102")
                {
                    if (dateExit > hourex1 && dateExit < hourex2)
                    {

                         worksheet.Cell(j, 8).Value = employees[i].DATE_TIME;

                        //уход на обед 
                    }
                    else
                    {
                         worksheet.Cell(j, 8).Value = "-";
                        //уход на обед пустой
                    }
                }
            }
            else
            {
                 worksheet.Cell(j, 8).Value = "--";
                //дата пустая
            }

        }
        public static string GetRes(string id)
        {
            DBResponse res = ListenTable(id);
            while (res.code == "0")
            {
                Thread.Sleep(1000);
                res = ListenTable(id);
            }
            return res.result;
        }
        public static DBResponse ListenTable(string id)
        {
            string constr = "DATA SOURCE=127.0.0.1:1521/orcl.local.com;PASSWORD=123;PERSIST SECURITY INFO=True;USER ID=SCU";// con sstring
            OracleConnection con = new OracleConnection();
            con.ConnectionString = constr;
            OracleCommand cmd = new OracleCommand("FLEX_GET_RESPONSE", con);
            cmd.Parameters.Add("ins_id", OracleDbType.Varchar2).Value = id;
            cmd.CommandType = CommandType.StoredProcedure;
            OracleParameter outparam = new OracleParameter();
            outparam = cmd.Parameters.Add("CURS", OracleDbType.RefCursor, ParameterDirection.Output);
            string result = "";
            string code = "0";
            try
            {
                con.Open();
                OracleDataReader rd = cmd.ExecuteReader();
                while (rd.Read())
                {
                    result = rd.GetValue(0).ToString();
                    code = rd.GetValue(1).ToString();
                }
                Console.WriteLine("добавлен файл");
                con.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                con.Close();
            }
            return new DBResponse(code, result);
        }


    }
}
