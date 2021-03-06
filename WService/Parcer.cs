﻿using Spire.Xls;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace WService
{
    class Parcer
    {
        string format = "dd.MM.yyyy h:mm:ss";
        string com = "", com2 = "";
        string[] prm = new string[10];
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;

        System.Globalization.CultureInfo provider =  System.Globalization.CultureInfo.InvariantCulture;

        System.Globalization.CultureInfo culture = System.Globalization.CultureInfo.InvariantCulture;

        Dictionary<string, List<string>> files = new Dictionary<string, List<string>>();
        Dictionary<string, string> formula = new Dictionary<string, string>();
        Dictionary<string, string> ranges = new Dictionary<string, string>();

        private static void ChangeStatus(string status, bool Increase)
        {

            MainWindow.StartWindow1.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
            {
                if(Increase==true)
                MainWindow.StartWindow1.pg1.Value = MainWindow.StartWindow1.pg1.Value + 1;

                MainWindow.StartWindow1.label1.Content = status;
            }));

        }             

        private void Files(List<string> commands)
        {
            
            foreach (var command in commands)
            {
                
                com = command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0];
                com2 = command.Split("@".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0];

                prm = (command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[1].Split(")".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0]).Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                if (com.Equals("Файл", StringComparison.InvariantCultureIgnoreCase) && prm[0] != "" && prm[0] != null && prm[1] != "" && prm[1] != null) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                    List<string> file_dir = new List<string>();

                    if (prm[2].Contains("xlsx"))
                    {
                        file_dir.Add(prm[1].TrimEnd().TrimStart());
                        file_dir.Add(prm[2].TrimEnd().TrimStart());
                    }
                    else
                    {
                        file_dir.Add(prm[1].TrimEnd().TrimStart());
                        file_dir.Add(GetFilename(prm[1].TrimEnd().TrimStart(), prm[2].TrimEnd().TrimStart())[0]);

                    }
                    
                    files.Add(prm[0], file_dir);

                }

                if (com2.Contains("Формула")) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                    formula.Add(com2.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[1].Split(")".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0], @command.Split("@".ToCharArray(), StringSplitOptions.None)[1]);

                }

                if (com.Equals("Диапазон", StringComparison.InvariantCultureIgnoreCase) && prm[0] != "" && prm[0] != null && prm[1] != "" && prm[1] != null) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                    List<string> file_dir = new List<string>();

                    if (files.TryGetValue(prm[1].TrimStart().TrimEnd(), out file_dir))
                    {

                        ranges.Add(prm[0], GetList(file_dir, Convert.ToInt32(prm[2])) + prm[3].TrimEnd().TrimStart());
                    }

                }


            }
        }                

        public Parcer(List<string> commands)
        {

        this.Files(commands);
         
        }

        static List<int> GetColumnByName(List<string> file_dir_2, string list, string name)
        {
            int list_num = 1;

            Int32.TryParse(list, out list_num);

            List<int> col = new List<int>();

            Workbook workbook = new Workbook();

            workbook.LoadFromFile(file_dir_2[0] + file_dir_2[1]);

            Worksheet sheet = workbook.Worksheets[list_num - 1];

            using (workbook)
            {
                using (sheet)
                {
                    CellRange[] ranges = sheet.FindAllString(name, false, false);
                    col.Add(ranges[0].Row); col.Add(ranges[0].Column);

                }

            }
            return col;

        }

        static List<int> GetColumnByName(string filename, string list, string name)
        {
            int list_num = 1;
            List<int> col = new List<int>();

            Int32.TryParse(list, out list_num);

            Workbook workbook = new Workbook();

            System.Globalization.CultureInfo cc = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.InvariantCulture;


            workbook.LoadFromFile(filename);

            Worksheet sheet = workbook.Worksheets[list_num - 1];

            using (workbook)
            {
                using (sheet)
                {
                    CellRange cr = sheet.FindString(name, true, true);

                    if (cr != null)
                    {
                        col.Add(sheet.FindString(name, true, true).Row);
                        col.Add(sheet.FindString(name, true, true).Column);
                    }

                }

            }
            return col;


        }

        static List<string> GetFilename(string file_directory, string str)
        {
            List<String> Files = new List<string>();

            Files = Directory.GetFiles(file_directory, "*.xlsx").ToList();

            List<string> ret = new List<string>();

            foreach (var file in Files)
            {

                if (GetColumnByName(file, "1", str).Count > 0)
                    ret.Add(file.ToString().Split("\\".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Last());


            }


            return ret;
        }

        static void ExelOpenSave(string filename)
        {
            try
            {
                MyApp = new Excel.Application();
                MyApp.Visible = true;

                MyBook = MyApp.Workbooks.Open(filename);

                MyBook.Save();
                MyApp.Quit();

            }
            finally
            {

                Marshal.ReleaseComObject(MyBook);
                Marshal.ReleaseComObject(MyApp);
                Marshal.FinalReleaseComObject(MyBook);
                Marshal.FinalReleaseComObject(MyApp);


                GC.Collect();
            }

        }

        private string GetList(List<string> file_dir, int list_num)
        {



            string name = "";

            Workbook workbook = new Workbook();

            workbook.LoadFromFile(file_dir[0] + file_dir[1]);

            Worksheet sheet = workbook.Worksheets[list_num - 1];

            using (workbook)
            {
                using (sheet)
                {
                    name = sheet.Name;
                }
            }

            return "'" + file_dir[0] + "[" + file_dir[1] + "]" + name + "'!";


        }

        static public void ScanScript(List<string> commands)
        {
            string com = "";
            string[] prm = new string[10];
            int total_count = 0;
            int available_count = 0;

            foreach (var command in commands)
            {
                com = command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0];

                if (com.Equals("Файл", StringComparison.InvariantCultureIgnoreCase)) 
                {
                    total_count++;
                }
            }

                MainWindow.StartWindow1.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
                {
                    MainWindow.StartWindow1.pg1.Minimum = 0;
                    MainWindow.StartWindow1.pg1.Maximum = total_count;
                }));

            foreach (var command in commands)
            {
                
                com = command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0];
               
                prm = (command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[1].Split(")".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0]).Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                               
                if (com.Equals("Файл", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                    MainWindow.StartWindow1.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
                    {
                        MainWindow.StartWindow1.pg1.Value = MainWindow.StartWindow1.pg1.Value + 1;
                        MainWindow.StartWindow1.label1.Content = "Проверяю " + command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + " " + prm[2];
                    }));
                                        

                    if (prm[2].Contains("xlsx"))
                    {
                        if (System.IO.File.Exists(prm[1]+ prm[2]))
                        {
                            available_count++;
                        }
                    }
                    else
                    {
                        if (GetFilename(prm[1].TrimEnd().TrimStart(), prm[2].TrimEnd().TrimStart()).Equals("")==false)
                         available_count++;                                            

                    }


                }
            }

            MainWindow.StartWindow1.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
            {              
                MainWindow.StartWindow1.label1.Content = "Доступно " + available_count + " из " + total_count;
            }));

        }

        public void ParceXL(List<string> commands)
            {
                        
            List<string> file_temp = new List<string>();
            string com = "", com2 = "";
            string[] prm = new string[10];
                    

            MainWindow.StartWindow1.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate ()
            {
                MainWindow.StartWindow1.pg1.Minimum = 0;
                    MainWindow.StartWindow1.pg1.Maximum = commands.Count();
            }));
                                 

            foreach (var command in commands)
            {

                
                ChangeStatus(command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0], true);

                com = command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0];
                com2 = command.Split("@".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0];

                prm = (command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[1].Split(")".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0]).Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);


                if (com.Equals("Файл", StringComparison.InvariantCultureIgnoreCase) == false && com.Equals("Формула", StringComparison.InvariantCultureIgnoreCase) == false && com.Equals("Диапазон", StringComparison.InvariantCultureIgnoreCase) == false)
                { 
                                 

                int worksheet_num = 0;
                int column_num = 1;
                int row_num = 1;

                string filename = prm[0];
                files.TryGetValue(prm[0], out file_temp);
                filename = file_temp[0] + file_temp[1];
                if (prm[1] != "" && prm[1] != null)
                    Int32.TryParse(prm[1], out worksheet_num);

                Workbook workbook = new Workbook();
                workbook.LoadFromFile(filename);
                Worksheet ws = workbook.Worksheets[worksheet_num - 1];

                int total_rows = ws.LastRow;
                int total_columns = ws.LastColumn;

                

                if (com.Equals("ВставитьФормулу", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {
                    
                    string formula_str = "";
                                        
                    formula.TryGetValue(prm[3], out formula_str);


                    if (prm[2] != "" && prm[2] != null)
                        Int32.TryParse(prm[2], out column_num);                    

                    var strings = ws.FindAllString(prm[2], false, false);

                    int first_row = strings.First().Row;
                    int first_column = strings.First().Column;

                    if (strings.Count() > 0)
                    {
                        for (int i = first_row + 1; i <= ws.LastRow; i++)
                        {
                            var currentFormula = "=" + formula_str;

                            ws.Range[i, first_column].Formula = currentFormula;

                            var formulaResult = workbook.CaculateFormulaValue(currentFormula);

                            var value = formulaResult.ToString();

                            ws.Range[i, first_column].Value = value;

                            ws.Range[i, first_column].Style = ws.Range[i, first_column + 1].Style;


                            Regex myReg = new Regex("[A-Z]\\d+");
                            MatchCollection matches = myReg.Matches(formula_str);
                            foreach (var match in matches)
                            {
                                Regex myReg2 = new Regex("\\d+");
                                int addr_number = Convert.ToInt32(myReg2.Match(match.ToString()).ToString()) + 1;
                                Regex myReg3 = new Regex("[A-Z]");
                                string addr_letter = myReg3.Match(match.ToString()).ToString();
                                formula_str = formula_str.Replace(match.ToString(), addr_letter + addr_number.ToString());
                                
                            }

                            ChangeStatus(command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + " (" + i.ToString() + " из " + ws.LastRow + " )", false);

                          
                        }


                    }

                    workbook.Save();

                    workbook.Dispose();


                }

                if (com.Equals("Склейка", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {


                    int second_book_position = 1;

                    string[] dir5 = Directory.GetFiles(@"C:\Обработки\Temp\", "*.xlsx");

                    Workbook workbook2 = new Workbook();

                    Worksheet ws2 = workbook2.Worksheets[0];
                    
                    workbook.LoadFromFile(dir5[0]);

                    ws = workbook.Worksheets[0];

                    var strings = ws.FindString(prm[0], false, false);

                    int first_row = strings.Row;

                    ws.Copy(ws.Rows[first_row], ws2.Range[second_book_position, 1], true);

                    second_book_position++;

                    workbook.Dispose();

                    int count = 0;

                    foreach (var _filename in dir5)
                    {
                        count++;

                        workbook = new Workbook();

                        workbook.LoadFromFile(_filename);

                        ws = workbook.Worksheets[0];

                        var range_to_copy = ws.Range[first_row + 2, ws.FirstColumn, ws.LastRow, ws.LastColumn];

                        ws.Copy(range_to_copy, ws2.Range[second_book_position, 1], true);

                        second_book_position = second_book_position + (ws.LastRow - first_row - 1);

                        workbook.Dispose();
                                               
                        ChangeStatus(command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + " (" + count.ToString() + " из " + dir5.Count() + " )", false);
                    }

                    workbook2.SaveToFile(prm[2] + ".xlsx", ExcelVersion.Version2010);

                    workbook2.Dispose();
                    


                }

                if (com.Equals("ВПР", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                                     
                    string column_name = "";

                    string range = "";                    

                    if (prm[2] != "" && prm[2] != null)
                        if (!Int32.TryParse(prm[2], out column_num))
                            column_name = prm[2];
                    

                    int strok_vniz = 0;


                    if (prm.Count() == 8)
                    {

                        range = prm[6];
                        ranges.TryGetValue(prm[6], out range);
                        strok_vniz = Convert.ToInt32(prm[3]);

                        for (int i = 1; i < ws.Rows.Count(); i++)
                        {
                            if (ws.GetCaculateValue(i, column_num).Equals("") == false)
                            {
                                row_num = i;
                                break;
                            }
                        }

                    }

                    if (prm.Count() == 7)
                    {
                        range = prm[5];
                        ranges.TryGetValue(prm[5], out range);
                        strok_vniz = Convert.ToInt32(prm[3]);

                        row_num = GetColumnByName(file_temp, prm[1], prm[2])[0];
                        column_num = GetColumnByName(file_temp, prm[1], prm[2])[1];


                    }

                    Workbook workbook2 = new Workbook();

                    workbook2.LoadFromFile((range.Split("'".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0].Split("[]".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + range.Split("'".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0].Split("[]".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[1]));

                    Worksheet ws2 = workbook2.Worksheets[range.Split("'".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0].Split("[]".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[2]];


                    var range_2 = ws2.Range[range.Split("!".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[1]];

                    int cells_in_table = ws.Rows.Count();

                    int offcet = Convert.ToInt32(prm[4]);

                    if (strok_vniz > 0)
                    {
                        for (int i = row_num + 1; i <= strok_vniz + row_num + 1; i++)
                        {
                            for (int j = range_2.Row; j <= range_2.LastRow; j++)
                            {
                                if (ws.GetCaculateValue(i, column_num + offcet).Equals(ws2.GetCaculateValue(j, range_2.Column)))
                                {
                                    ws.SetCellValue(i, column_num, ws2.GetCaculateValue(j, range_2.Row + Convert.ToInt32(prm[6]) - 1).ToString());
                                    ws.Range[i, column_num].Style = ws.Range[i, column_num + 1].Style;
                                }
                               
                            }
                        }
                    }

                    if (strok_vniz == 0)
                    {
                        for (int i = row_num + 1; i <= cells_in_table; i++)
                        {


                            for (int j = range_2.Row; j <= ws2.Columns[0].RowCount; j++)
                            {

                                if (ws.GetCaculateValue(i, column_num + offcet).Equals(ws2.GetCaculateValue(j, range_2.Column)) && ws.GetCaculateValue(i, column_num + offcet).Equals("") == false)
                                {
                                    ws.SetCellValue(i, column_num, ws2.GetCaculateValue(j, range_2.Column + Convert.ToInt32(prm[6]) - 1).ToString());
                                    ws.Range[i, column_num].Style = ws.Range[i, column_num - 1].Style;
                                    break;
                                }
                                

                            }

                            ChangeStatus(command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + " (" + i.ToString() + " из " + cells_in_table + " )", false);
                            
                        }

                        workbook2.Dispose();
                    }


                    workbook.Save();
                    workbook.Dispose();
                                       

                }

                if (com.Equals("УдалитьПоДату", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                    string worksheet_name = "";
                   
                    files.TryGetValue(prm[0], out file_temp);
                    filename = file_temp[0] + file_temp[1];

                    if (prm[1] != "" && prm[1] != null)
                    {
                        if (!Int32.TryParse(prm[1], out worksheet_num))
                            worksheet_name = prm[1];

                    }
                    

                    int first = ws.FirstVisibleRow;

                     row_num = GetColumnByName(file_temp, prm[1], prm[2])[0];
                     column_num = GetColumnByName(file_temp, prm[1], prm[2])[1];

                    DateTime user_date = DateTime.ParseExact(prm[3] + " 00:00:00", format, provider);

                    DateTime table_date;
                    
                    first = ws.FirstVisibleRow;

                    row_num = GetColumnByName(file_temp, prm[1], prm[2])[0] - 1;
                    column_num = GetColumnByName(file_temp, prm[1], prm[2])[1];

                    var format1 = "MM/dd/yyyy 00:00:00";

                    for (int i = row_num + 2; i < ws.LastRow; i++)
                    {
                        if (ws.GetCaculateValue(i, column_num).ToString().Equals(""))
                        {
                            ws.DeleteRow(i);
                            i--;
                        }

                        table_date = DateTime.ParseExact(ws.GetCaculateValue(i, column_num).ToString().TrimEnd().TrimStart(), format1, provider);

                        if (table_date >= user_date)
                        {

                            ws.DeleteRow(i);
                            i--;
                        }

                        ChangeStatus(command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + " (" + i.ToString() + " из " + ws.LastRow + " )", false);
                                               

                    }


                    workbook.Save();


                }

                if (com.Equals("УдалитьНоль", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                               
                  
                    if (prm[2] != "" && prm[2] != null)
                        Int32.TryParse(prm[2], out column_num);

                    var strings = ws.FindAllString(prm[2], false, false);

                    if (strings.Count() > 0)
                    {
                        var column = strings[0].Column;
                        var row = strings[0].Row + 1;
                        var range = ws.Range[row, column];

                        int last_row = ws.LastRow;


                        for (int j = row; j <= ws.LastRow; j++)
                        {

                            if (ws.GetCaculateValue(j, column).ToString().Equals("0"))
                            {
                                ws.DeleteRow(j);
                                j--;

                            }

                            ChangeStatus(command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + " (" + j.ToString() + " из " + ws.LastRow + " )", false);

                            
                        }


                    }
                    
                    workbook.Save();

                    workbook.Dispose();

                }

                if (com.Equals("УдалитьУволен", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                    if (prm[2] != "" && prm[2] != null)
                        Int32.TryParse(prm[2], out column_num);
                    
                    var strings = ws.FindAllString("✘", false, false);
                    int i = 0;
                    foreach (var str in strings)
                    {
                        ws.DeleteRow(str.Row - i);
                        i++;

                        ChangeStatus(command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + " (" + i.ToString() + " из " + strings.Count() + " )", false);                                             

                    }

                    workbook.Save();

                    workbook.Dispose();
                    
                }

                if (com.Equals("УдалитьСтолбец", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                                      
                    if (prm[2] != "" && prm[2] != null)
                        Int32.TryParse(prm[2], out column_num);

                    
                    var strings = ws.FindAllString(prm[2], false, false);

                    if (strings.Count() > 0)
                        ws.DeleteColumn(ws.FindAllString(prm[2], false, false)[0].Column);

                    workbook.Save();

                    workbook.Dispose();
                    
                }

                if (com.Equals("ДобавитьСтроку", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                    string column_name = "";

                
                    if (prm[2] != "" && prm[2] != null)
                    {

                        if (!Int32.TryParse(prm[2], out column_num))
                            column_name = prm[2];

                    }
                   
                    if (column_name != "")
                        column_num = GetColumnByName(file_temp, prm[1], prm[2])[1];

                    for (int i = 1; i < ws.LastRow; i++)
                    {
                        if (ws.GetCaculateValue(i, column_num).Equals("") == false)
                        {
                            row_num = i;
                            break;
                        }

                    }


                    ws.InsertColumn(column_num);
                    ws.SetCellValue(row_num, column_num, prm[3].ToString());

                    ws.Range[row_num, column_num].Style = ws.Range[row_num, column_num - 1].Style;


                    workbook.Save(); workbook.Dispose();
                    
                }

                if (com.Equals("ПоСодержанию", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                    files.TryGetValue(prm[0], out file_temp);
                    filename = file_temp[0] + file_temp[1];


                    if (prm[1] != "" && prm[1] != null)
                        Int32.TryParse(prm[1], out worksheet_num);

                    List<string> clients = new List<string>();

                    if (prm[2].Contains(";"))
                        clients.AddRange(prm[2].Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries));
                    else
                        clients.Add(prm[2]);



                    Workbook workbook2 = new Workbook();



                    Worksheet ws2 = workbook2.Worksheets[0];
                    workbook2.Worksheets[1].Remove();
                    workbook2.Worksheets[1].Remove();

                    int first = ws.FirstVisibleRow;

                    row_num = GetColumnByName(file_temp, prm[1], prm[3])[0] - 1;
                    column_num = GetColumnByName(file_temp, prm[1], prm[3])[1];

                    int new_row = 1;

                    ws.Copy(ws.Rows[row_num], ws2.Range[new_row, ws.FirstColumn], true);
                    new_row++;


                    for (int i = row_num + 2; i <= total_rows; i++)
                    {

                        ChangeStatus(command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + " (" + i.ToString() + " из " + total_rows + " )", false);
                    
                        foreach (var client in clients)
                        {
                            if (client.Contains("!"))
                            {
                                if (ws.GetCaculateValue(i, column_num).ToString().TrimEnd().TrimStart().ToUpper().Contains(client.ToString().Replace("!", "").TrimEnd().TrimStart().ToUpper()) == false)
                                {

                                    ws.Copy(ws.Rows[i - 1], ws2.Range[new_row, ws.FirstColumn], true);

                                    new_row++;

                                }
                            }
                            else
                            {
                                if (ws.GetCaculateValue(i, column_num).ToString().TrimEnd().TrimStart().ToUpper().Contains(client.ToString().TrimEnd().TrimStart().ToUpper()) == true)
                                {

                                    ws.Copy(ws.Rows[i - 1], ws2.Range[new_row, ws.FirstColumn], true);

                                    new_row++;
                                }
                            }

                        }

                    }

                    if (prm[5] != null)

                    column_num = GetColumnByName(file_temp, prm[1], prm[5])[1];

                    workbook2.DataSorter.SortColumns.Add(column_num - 1, OrderBy.Ascending);

                    var range = ws2.Range[ws2.FirstRow, ws2.FirstColumn, ws2.LastRow, ws2.LastColumn];

                    workbook2.DataSorter.Sort(range);

                    workbook2.SaveToFile(prm[4] + ".xlsx", ExcelVersion.Version2010);

                    workbook2.Dispose(); workbook.Dispose();
                   
                }

                if (com.Equals("ДобавитьДокументы", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                                     
                    string column_name = "";
                                     

                    if (prm[2] != "" && prm[2] != null)
                    {

                        if (!Int32.TryParse(prm[2], out column_num))
                            column_name = prm[2];

                    }
                    

                    if (column_name != "")
                        column_num = GetColumnByName(file_temp, prm[1], prm[2])[1];

                  
                    for (int i = 1; i < total_rows; i++)
                    {
                        if (ws.GetCaculateValue(i, column_num).Equals("") == false)
                        {
                            row_num = i;
                            break;
                        }

                        ChangeStatus(command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + " (" + i.ToString() + " из " + total_rows + " )", false);
                                               
                    }


                    ws.InsertColumn(column_num);
                    ws.SetCellValue(row_num, column_num, prm[3].ToString());
                    ws.Range[row_num, column_num].Style = ws.Range[row_num, column_num + 1].Style;

                    for (int i = row_num + 1; i < total_rows; i++)
                    {
                        if (ws.GetCaculateValue(i, column_num + 1).Equals("") == true && ws.GetCaculateValue(i, column_num + 2).Equals("") == true)
                        {

                            ws.SetCellValue(i, column_num, "без документов");
                            ws.Range[i, column_num].Style = ws.Range[i, column_num + 1].Style;

                        }
                        if (ws.GetCaculateValue(i, column_num + 1).Equals("") == false || ws.GetCaculateValue(i, column_num + 2).Equals("") == false)

                        {
                            ws.SetCellValue(i, column_num, "с документами");
                            ws.Range[i, column_num].Style = ws.Range[i, column_num + 1].Style;

                        }

                        ChangeStatus(command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + " (" + i.ToString() + " из " + total_rows + " )", false);
                                               
                    }


                    workbook.Save(); workbook.Dispose();
                                        
                }

                if (com.Equals("ОбработатьПланы", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                    string worksheet_name = "";
                   
                    if (prm[1] != "" && prm[1] != null)
                    {

                        if (!Int32.TryParse(prm[1], out worksheet_num))
                            worksheet_name = prm[1];

                    }
                                      
                    Worksheet ws2 = workbook.CreateEmptySheet();                                    

                    int new_row = 1;



                    ws2.SetValue(new_row, 1, "ФИО");
                    ws2.SetValue(new_row, 2, "Филиал");
                    ws2.SetValue(new_row, 3, "Ответственный");
                    ws2.SetValue(new_row, 16, "Дней факт (прошлая неделя)");
                    ws2.SetValue(new_row, 17, "План 2 дня (текущая неделя)");

                    ws2.SetValue(new_row, 18, "Санкции факт");
                    ws2.SetValue(new_row, 19, "Санкции план");
                    ws2.SetValue(new_row, 20, "Итого");

                    ws2.Name = "Планирование_обр";
                    ws.Name = "Планирование_оракл";

                    for (int j = 9; j <= 20; j++)
                    {
                        ws2.SetValue(new_row, j - 5, ws.GetCaculateValue(4, j).ToString() + ws.GetCaculateValue(5, j).ToString());

                    }


                    new_row++;

                    for (int i = 6; i <= total_rows; i++)
                    {

                        ChangeStatus(command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + " (" + i.ToString() + " из " + total_rows + " )", false);
                        

                        if (ws.GetCaculateValue(i, 2).ToString().Contains("осн/рег/all") && ws.Range[i, 9].Style.KnownColor != ws.Range[i + 1, 9].Style.KnownColor)
                        {

                            int sum_plan = 0, sum_fact = 0;

                            ws2.SetValue(new_row, 1, ws.GetCaculateValue(i, 2).ToString().Replace("осн/рег/all", ""));
                            ws2.SetValue(new_row, 2, ws.GetCaculateValue(i, 8).ToString());

                            for (int j = 9; j <= 20; j++)
                            {
                                ws2.SetValue(new_row, j - 5, ws.GetCaculateValue(i, j).ToString());

                                ws2.Range[new_row, j - 5].Style = ws.Range[i, j].Style;


                                if (ws.GetCaculateValue(i, j).ToString().Equals("") == false)
                                {
                                    if (ws.GetCaculateValue(i, j).ToString().Contains("/") || ((j <= 13 && ws.Range[i, j].Style.KnownColor.ToString().Contains("LightOrange")) || (j <= 13 && ws.Range[i, j].Style.KnownColor.ToString().Contains("LightYellow"))))
                                    {
                                        if (j <= 13)
                                            sum_fact++;

                                        if (j >= 19 && ws.GetCaculateValue(i, j).ToString().Split("/".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0].Equals("") || ws.GetCaculateValue(i, j).ToString().Split("/".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0].Equals(" "))
                                            sum_plan++;

                                    }
                                    else
                                    {
                                        if (j >= 19)
                                            sum_plan++;

                                    }


                                }
                                else
                                {

                                    if ((j <= 13 && ws.Range[i, j].Style.KnownColor.ToString().Contains("LightOrange")) || (j <= 13 && ws.Range[i, j].Style.KnownColor.ToString().Contains("LightYellow")))
                                        sum_fact++;
                                    if ((j >= 19 && ws.Range[i, j].Style.KnownColor.ToString().Contains("LightOrange")) || (j >= 19 && ws.Range[i, j].Style.KnownColor.ToString().Contains("LightYellow")))
                                        sum_plan++;


                                }

                                ws2.SetValue(new_row, 16, sum_fact.ToString());
                                ws2.SetValue(new_row, 17, sum_plan.ToString());

                                // =ЕСЛИ(P2<5;300;0)

                                ws2[new_row, 18].Formula = "=IF(" + ws2.Range[new_row, 16].RangeAddressLocal.ToString() + "<5,300,0)";
                                ws2[new_row, 19].Formula = "=IF(" + ws2.Range[new_row, 17].RangeAddressLocal.ToString() + "<2,100,0)";
                                ws2[new_row, 20].Formula = "=" + ws2.Range[new_row, 18].RangeAddressLocal.ToString() + "+" + ws2.Range[new_row, 19].RangeAddressLocal.ToString();




                            }

                            new_row++;


                        }


                    }

                    for (int j = 1; j <= ws2.Columns.Count(); j++)
                    {

                        ws2.Range[1, j].Style.KnownColor = ExcelColors.Gray25Percent;
                        ws2.Range[1, j].Style.ShrinkToFit = true;
                        ws2.Range[1, j].Style.HorizontalAlignment = HorizontalAlignType.Center;
                        ws2.Range[1, j].Style.VerticalAlignment = VerticalAlignType.Center;
                        ws2.Range[1, j].Style.WrapText = true;
                        ws2.Range[1, j].Style.Font.IsBold = true;
                    }

                    ws2.Range[1, ws2.LastColumn, ws2.LastRow, ws2.LastColumn].AutoFitColumns();
                    ws2.Range[1, ws2.LastColumn, ws2.LastRow, ws2.LastColumn].AutoFitRows();

                    workbook.Save(); workbook.Dispose();
                    
                }

                if (com.Equals("ОбработатьККД2", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                    string worksheet_name = "";
                    

                    if (prm[1] != "" && prm[1] != null)
                    {

                        if (!Int32.TryParse(prm[1], out worksheet_num))
                            worksheet_name = prm[1];

                    }                 

                    for (int i = 1; i < total_rows; i++)
                    {
                       
                       ChangeStatus(command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + " (" + i.ToString() + " из " + total_rows + " )", false);
                        

                        for (int j = 1; j < total_columns; j++)
                        {
                            if (ws.GetCaculateValue(i, j).ToString().Contains("Великий Новгород"))
                                ws.SetValue(i, j, "Великий-Новгород");

                            if (ws.GetCaculateValue(i, j).ToString().Contains("%") == true)
                            {

                                if (ws.GetCaculateValue(i, j).ToString().Contains("/"))
                                {
                                    if (ws.GetCaculateValue(i, j).ToString().Split("/".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Count() > 1)
                                        ws.SetValue(i, j, ws.GetCaculateValue(i, j).ToString().Split("/".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[1].Split(" ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0].TrimEnd().TrimStart());
                                    else
                                        ws.SetValue(i, j, ws.GetCaculateValue(i, j).ToString().Split("/".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0].Split(" ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0].TrimEnd().TrimStart());

                                }
                            }
                        }
                    }


                    workbook.Save(); workbook.Dispose();
                    
                }

                if (com.Equals("ОбработатьККД", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {
                    
                    for (int i = 1; i < total_rows; i++)
                    {

                        ChangeStatus(command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + " (" + i.ToString() + " из " + total_rows + " )", false);
                      
                        for (int j = 1; j < total_columns; j++)
                        {
                            if (ws.GetCaculateValue(i, j).ToString().Contains("Великий Новгород"))
                                ws.SetValue(i, j, "Великий-Новгород");

                            if (ws.GetCaculateValue(i, j).ToString().Contains("%") == true)
                            {

                                if (ws.GetCaculateValue(i, j).ToString().Contains("/"))
                                {
                                    if (ws.GetCaculateValue(i, j).ToString().Split("/".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Count() > 1)
                                    {

                                        ws.SetValue(i, j, ws.GetCaculateValue(i, j).ToString().Split("/".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[1].Split(" ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0].TrimEnd().TrimStart().Replace(",", "."));

                                    }
                                    else
                                        ws.SetValue(i, j, ws.GetCaculateValue(i, j).ToString().Split("/".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0].Split(" ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0].TrimEnd().TrimStart());

                                }
                            }
                        }
                    }


                    workbook.Save(); workbook.Dispose();
                                        
                }
                if (com.Equals("ОбработатьМАКС", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {
                                      
                    bool control = false;

                    foreach (var worksheet in workbook.Worksheets)
                    {

                        if (worksheet.Name.Contains("Сводная"))
                            control = true;

                    }

                    if (control == false)
                    {

                        for (int i = 1; i < total_rows; i++)
                        {
                            ChangeStatus(command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + " (" + i.ToString() + " из " + total_rows + " )", false);

                           
                            for (int j = 1; j < total_columns; j++)
                            {

                                if (ws.GetCaculateValue(i, j).ToString().Equals("Сумма") || ws.GetCaculateValue(i, j).ToString().Equals("Итого") || ws.GetCaculateValue(i, j).ToString().Equals("Подч.") || ws.GetCaculateValue(i, j).ToString().Equals("Перс.") || ws.GetCaculateValue(i, j).ToString().Equals("Кол-во"))
                                    ws.SetValue(i, j, ws.GetCaculateValue(i, j) + j.ToString());
                                if (ws.GetCaculateValue(i, j).ToString().Equals("Пл/Фт"))
                                    ws.DeleteColumn(j);

                            }
                        }

                        workbook.SaveToFile(filename.Split(".".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + "_обр" + ".xlsx", ExcelVersion.Version2010);

                        if (File.Exists(filename))
                            File.Delete(filename);

                        File.Move(filename.Split(".".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + "_обр" + ".xlsx", filename);

                        workbook.Dispose();



                    }

                }

                if (com.Equals("ОбработатьЗО", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                    foreach (var worksheet in workbook.Worksheets)
                    {

                        worksheet.Name = worksheet.Name.Replace("(", "_");
                        worksheet.Name = worksheet.Name.Replace(")", "_");


                    }

                    for (int i = 1; i < total_rows; i++)
                    {
                        ChangeStatus(command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + " (" + i.ToString() + " из " + total_rows + " )", false);
                                              

                        for (int j = 1; j < total_columns; j++)
                        {

                            if (i < 5 && ws.GetCaculateValue(i, j).ToString().Contains("("))
                                ws.SetValue(i, j, ws.GetCaculateValue(i, j).ToString().Replace("(", "").Replace(")", ""));

                            if (ws.GetCaculateValue(i, j).ToString().Contains("В.Новгород"))
                                ws.SetValue(i, j, "Великий-Новгород");
                            if (ws.GetCaculateValue(i, j).ToString().Contains("Хабаровск-Кондер"))
                                ws.SetValue(i, j, "Хабаровск");

                            if (ws.GetCaculateValue(i, j).ToString().Contains("Сервис "))
                                ws.SetValue(i, j, ws.GetCaculateValue(i, j).ToString().Replace("Сервис ", ""));
                            if (ws.GetCaculateValue(i, j).ToString().Contains("Аренда "))
                                ws.SetValue(i, j, ws.GetCaculateValue(i, j).ToString().Replace("Аренда ", ""));

                        }


                    }
                                        
                    workbook.SaveToFile(filename.Split(".".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + "_обр" + ".xlsx");

                    if (File.Exists(filename))
                        File.Delete(filename);

                    File.Move(filename.Split(".".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + "_обр" + ".xlsx", filename);

                    workbook.Dispose();                   

                }

                if (com.Equals("ОбработатьДебиторка", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                  
                    for (int i = 1; i < total_rows; i++)
                    {
                        ChangeStatus(command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + " (" + i.ToString() + " из " + total_rows + " )", false);

                        for (int j = 1; j < total_columns; j++)
                        {

                            if (ws.GetCaculateValue(i, j).ToString().Contains("В.Новгород"))
                                ws.SetValue(i, j, "Великий-Новгород");
                            if (ws.GetCaculateValue(i, j).ToString().Contains("Хабаровск-Кондер"))
                                ws.SetValue(i, j, "Хабаровск");
                            
                        }
                    }


                    int first = ws.FirstVisibleRow;

                    row_num = GetColumnByName(file_temp, prm[1], prm[2])[0];
                    column_num = GetColumnByName(file_temp, prm[1], prm[2])[1];

                    DateTime user_date = DateTime.ParseExact(prm[3] + " 00:00:00", format, provider);
                    DateTime start_date = user_date.AddYears(-3);
                    DateTime table_date;

                    int new_row = 1;

                    if (File.Exists(file_temp[0] + prm[4] + ".xlsx"))
                        File.Delete(file_temp[0] + prm[4] + ".xlsx");

                    Workbook workbook2 = new Workbook();


                    Worksheet ws2 = workbook2.Worksheets[0];
                    workbook2.Worksheets[1].Remove();
                    workbook2.Worksheets[1].Remove();

                    total_rows = ws.Rows.Count();
                    total_columns = ws.Columns.Count();
                    first = ws.FirstVisibleRow;

                    row_num = GetColumnByName(file_temp, prm[1], prm[2])[0] - 1;
                    column_num = GetColumnByName(file_temp, prm[1], prm[2])[1];

                    new_row = 1;

                    ws.Copy(ws.Rows[row_num], ws2.Range[new_row, ws.FirstColumn], true);

                    var format1 = "MM/dd/yyyy 00:00:00";

                    for (int i = row_num + 2; i < total_rows; i++)
                    {
                       
                            table_date = DateTime.ParseExact(ws.GetCaculateValue(i, column_num).ToString().TrimEnd().TrimStart(), format1, provider);

                            if (table_date >= start_date && table_date < user_date)
                            {
                                new_row++;
                                
                                ws.Copy(ws.Rows[i - 1], ws2.Range[new_row, ws.FirstColumn], true);

                            }

                        ChangeStatus(command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + " (" + i.ToString() + " из " + total_rows + " )", false);
                                              
                    }
                                       

                    if (prm[4].Contains("\\") == false)
                        workbook2.SaveToFile(file_temp[0] + prm[4] + ".xlsx", ExcelVersion.Version2010);
                    else
                        workbook2.SaveToFile(prm[4] + ".xlsx", ExcelVersion.Version2010);

                    workbook2.Dispose(); workbook.Dispose();
                  

                }

                if (com.Equals("ПоКлиенту", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {

                  
                    List<string> clients = new List<string>();

                    if (prm[2].Contains(";"))
                        clients.AddRange(prm[2].Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries));
                    else
                        clients.Add(prm[2]);

                    Workbook workbook2 = new Workbook();


                    Worksheet ws2 = workbook2.Worksheets[0];
                    workbook2.Worksheets[1].Remove();
                    workbook2.Worksheets[1].Remove();
                                      
                    int first = ws.FirstVisibleRow;

                    row_num = GetColumnByName(file_temp, prm[1], prm[3])[0] - 1;
                    column_num = GetColumnByName(file_temp, prm[1], prm[3])[1];

                    int new_row = 1;

                    ws.Copy(ws.Rows[row_num], ws2.Range[new_row, ws.FirstColumn], true);
                    new_row++;


                    for (int i = row_num + 2; i <= total_rows; i++)
                    {
                        ChangeStatus(command.Split("(".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + " (" + i.ToString() + " из " + total_rows + " )", false);
                        
                        foreach (var client in clients)
                        {
                            if (client.Contains("!"))
                            {
                                if (ws.GetCaculateValue(i, column_num).ToString().TrimEnd().TrimStart().ToUpper().Equals(client.ToString().Replace("!", "").TrimEnd().TrimStart().ToUpper()) == false)
                                {

                                    ws.Copy(ws.Rows[i - 1], ws2.Range[new_row, ws.FirstColumn], true);

                                    new_row++;

                                }
                            }
                            else
                            {
                                if (ws.GetCaculateValue(i, column_num).ToString().TrimEnd().TrimStart().ToUpper().Equals(client.ToString().TrimEnd().TrimStart().ToUpper()) == true)
                                {

                                    ws.Copy(ws.Rows[i - 1], ws2.Range[new_row, ws.FirstColumn], true);

                                    new_row++;
                                }
                            }

                        }

                    }

                    if(prm[4].Contains("\\")==false)
                    workbook2.SaveToFile(file_temp[0] + prm[4] + ".xlsx", ExcelVersion.Version2010);
                    else
                    workbook2.SaveToFile(prm[4] + ".xlsx", ExcelVersion.Version2010);

                    workbook2.Dispose(); workbook.Dispose();
                    
                }

                if (com.Equals("Сводная", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {
                                       
                    bool control = false;
                                      
                    column_num = GetColumnByName(file_temp, prm[1], prm[5].Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0])[1];

                    for (int i = ws.FirstRow; i < ws.LastRow; i++)
                        if (ws.GetCaculateValue(i, column_num).ToString().Equals(""))
                            ws.SetNumber(i, column_num, 0);

                    workbook.Save();
                    workbook.Dispose();

                    ExelOpenSave(filename);

                    workbook = new Workbook();

                    workbook.LoadFromFile(filename);

                    ws = workbook.Worksheets[worksheet_num - 1];

                    List<string> RowLabels = new List<string>(), ColumnLabels = new List<string>();

                    if (prm[3].Contains(";"))
                    {

                        RowLabels.AddRange(prm[3].Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries));
                    }
                    else
                    {
                        RowLabels.Add(prm[3]);
                    }
                    if (prm[4].Contains(";"))
                    {

                        ColumnLabels.AddRange(prm[4].Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries));
                    }
                    else
                    {
                        ColumnLabels.Add(prm[4]);
                    }




                    foreach (var worksheet in workbook.Worksheets)
                    {

                        if (worksheet.Name.Contains("Сводная"))
                            control = true;


                    }

                    if (control == false)
                    {



                        int last_row_used = ws.LastRow - 1;

                        ws.Name = ws.Name.ToString().Replace("(", "").Replace(")", "");

                        int first_table_row = GetColumnByName(file_temp, prm[1], RowLabels[0])[0]; //вычисляем порядковый номер первой строки таблицы

                        var source = ws.Range[first_table_row, 1, last_row_used - 1, ws.LastColumn]; // формируем диапазон сводной талицы

                        PivotCache cache = workbook.PivotCaches.Add(source);

                        var wsPT = workbook.Worksheets.Create(prm[2]);

                        PivotTable pt = wsPT.PivotTables.Add(prm[2], ws.Range["A1"], cache);




                        foreach (var rowlabel in RowLabels)
                        {
                            var r1 = pt.PivotFields[rowlabel];

                            r1.Axis = AxisTypes.Row;

                            pt.Options.RowHeaderCaption = rowlabel;
                        }

                        if (prm[5].Contains(";"))
                        {

                            RowLabels.AddRange(prm[5].Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries));



                            pt.DataFields.Add(pt.PivotFields[prm[5].Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0]], prm[5].Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0], SubtotalTypes.Sum);



                            pt.DataFields.Add(pt.PivotFields[prm[5].Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[1]], prm[5].Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[1], SubtotalTypes.Count);

                        }
                        else
                        {
                            if (prm[5] != "" && prm[5] != " ")
                                pt.DataFields.Add(pt.PivotFields[prm[5]], prm[5], SubtotalTypes.Sum);
                        }


                        workbook.SaveToFile(filename.Split(".".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + "_обр" + ".xlsx");

                        if (File.Exists(filename))
                            File.Delete(filename);

                        File.Move(filename.Split(".".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)[0] + "_обр" + ".xlsx", filename);

                        workbook.Dispose();

                        ExelOpenSave(filename);


                    }
                    

                }

                if (com.Equals("Разбить", StringComparison.InvariantCultureIgnoreCase)) // добавляем строку до столбца по номеру или до слолбца по названию
                {
              

                    List<string> clients = new List<string>();
                    if (prm[2].Contains(";"))
                        clients.AddRange(prm[2].Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries));
                    else
                        clients.Add(prm[2]);
                   

                    
                    int first = ws.FirstVisibleRow;
                    row_num = GetColumnByName(file_temp, prm[1], prm[2])[0] - 1;
                    column_num = GetColumnByName(file_temp, prm[1], prm[2])[1];
                    int new_row = 1;

                    bool series = false;
                    Workbook workbook2 = new Workbook();
                    Worksheet ws2 = workbook2.CreateEmptySheet("Мотивация куратора");

                    for (int i = row_num + 2; i < total_rows; i++)
                    {

                        if (ws.GetCaculateValue(i, column_num).ToString().TrimEnd().TrimStart().ToUpper().Equals(ws.GetCaculateValue(i - 1, column_num).ToString().TrimEnd().TrimStart().ToUpper()) == true && series == false)
                        {
                            workbook2 = new Workbook();
                            ws2 = workbook2.Worksheets[0];
                            ws2.Name = ws.Name;
                            new_row = 1;
                            ws.Copy(ws.Rows[row_num], ws2.Range[new_row, 1, new_row, total_columns], true);
                            new_row++;
                            ws.Copy(ws.Rows[i - 2], ws2.Range[new_row, 1, new_row, total_columns], true);
                            new_row++;
                            ws.Copy(ws.Rows[i - 1], ws2.Range[new_row, 1, new_row, total_columns], true);
                            new_row++;
                            i++;
                            series = true;
                        }
                        if (ws.GetCaculateValue(i, column_num).ToString().TrimEnd().TrimStart().ToUpper().Equals(ws.GetCaculateValue(i - 1, column_num).ToString().TrimEnd().TrimStart().ToUpper()) == true && series == true)
                        {
                            ws.Copy(ws.Rows[i - 1], ws2.Range[new_row, 1, new_row, total_columns], true);

                            new_row++;
                        }
                        if (ws.GetCaculateValue(i, column_num).ToString().TrimEnd().TrimStart().ToUpper().Equals(ws.GetCaculateValue(i - 1, column_num).ToString().TrimEnd().TrimStart().ToUpper()) == false && series == true)
                        {

                            for (int j = 1; j <= ws.LastColumn; j++)
                            {
                                string formula_str = ws.GetCaculateValue(i, j).ToString();                                                              

                                if (new_row < i)
                                { 
                                    Regex myReg = new Regex("[A-Z]\\d+");
                                    MatchCollection matches = myReg.Matches(formula_str);
                                   
                                    foreach (var match in matches)
                                    {
                                        Regex myReg2 = new Regex("\\d+");
                                        int addr_number = Convert.ToInt32(myReg2.Match(match.ToString()).ToString()) - (i-new_row);
                                        Regex myReg3 = new Regex("[A-Z]");
                                        string addr_letter = myReg3.Match(match.ToString()).ToString();
                                        formula_str = formula_str.Replace(match.ToString(), addr_letter + addr_number.ToString());
                                    

                                    }
                                }
                                ws2.SetValue(new_row, j, formula_str);

                                ws2.Range[new_row, j].Style.KnownColor = ws.Range[i, j].Style.KnownColor;

                                if (ws.Range[i, j].Style.NumberFormat != null)
                                    ws2.Range[new_row, j].Style.NumberFormat = ws.Range[i, j].Style.NumberFormat;
                            }
                            

                            workbook2.SaveToFile(prm[3] + ws.GetCaculateValue(i - 1, column_num).ToString().Replace(".", "").Replace(" ", "") + ".xlsx", ExcelVersion.Version2010);
                            series = false;
                            workbook2.Dispose();
                        }
                    }
                                        
                }


            }

            ChangeStatus("Все операции завершились успешно.", false);
            

            }
        }

        static public string ExcelSave2010(string filename, string save_to)
        {


            try
            {
                MyApp = new Excel.Application();
                MyApp.Visible = true;


                MyBook = MyApp.Workbooks.Open(filename);
                string new_file = filename.Split("\\".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Last().Split(".".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).First() + ".xlsx";
                MyBook.SaveAs(save_to + new_file, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                MyApp.Quit();
                return new_file;

            }

            finally
            {

                Marshal.ReleaseComObject(MyBook);
                Marshal.ReleaseComObject(MyApp);
                Marshal.FinalReleaseComObject(MyBook);
                Marshal.FinalReleaseComObject(MyApp);

                GC.Collect();
            }



        }
    }
}
