﻿using Spire.Xls;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Windows.Threading;

namespace WService
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    /// 
    

    public partial class MainWindow : Window
    {



        public static MainWindow StartWindow1;

        public MainWindow()
        {

                      
            InitializeComponent();

            StartWindow1 = this;
                        

            string[] dirs3 = Directory.GetFiles(@"C:\Обработки\Temp\", "*.xlsx");
            listBox2.Items.Clear();
            listBox2.ItemsSource = dirs3;

            string[] dirs = Directory.GetFiles("c:\\XLTest\\", "*.txt");
            listBox.Items.Clear();
            listBox.ItemsSource = dirs;
            
            string[] dirs2 = Directory.GetFiles(@"C:\Users\kozlov.r\Downloads", "*.xls");
            
            listBox1.Items.Clear();
            listBox1.ItemsSource = dirs2.OrderByDescending(s => new FileInfo(s).CreationTime ); 

            

        }

        private void listBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            label.Content = "Информация о создании файла: " + new FileInfo(listBox1.SelectedItem.ToString()).CreationTime + " Размер: " + (Convert.ToSingle(new FileInfo(listBox1.SelectedItem.ToString()).Length)/1024/1024).ToString("#.##") + " Мб.";
        }

        private void listBox1_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Parcer.ExcelSave2010(listBox1.SelectedItem.ToString(), @"C:\Обработки\Temp\");
            string[] dirs3 = Directory.GetFiles(@"C:\Обработки\Temp\", "*.xlsx");
            listBox2.ItemsSource = dirs3;
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {

           
            foreach (var filename in Directory.GetFiles(@"C:\Обработки\Temp\", "*.xlsx"))
               File.Delete(filename);

            string[] dirs3 = Directory.GetFiles(@"C:\Обработки\Temp\", "*.xlsx");
            listBox2.ItemsSource = dirs3;


        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.InvariantCulture;

            List<string> lst = new List<string>();
            string line;
            lst.Clear();

            // Read the file and display it line by line.  
            System.IO.StreamReader file =
                new System.IO.StreamReader(listBox.SelectedItem.ToString(), Encoding.GetEncoding(1251));

            while ((line = file.ReadLine()) != null)
                lst.Add(line);

            file.Close();

            Parcer parcer = new Parcer();

            var task = new Task(new Action(() => parcer.ParceXL(lst)));

            task.Start();

           

        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {

           

        }

        private void listBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.InvariantCulture;

            List<string> lst = new List<string>();
            string line;
            
            lst.Clear();
            

                // Read the file and display it line by line.  
                System.IO.StreamReader file =
                    new System.IO.StreamReader(listBox.SelectedItem.ToString(), Encoding.GetEncoding(1251));

                while ((line = file.ReadLine()) != null)
                    lst.Add(line);


            
            var task = new Task(new Action(() => Parcer.ScanScript(lst)));

            task.Start();



            file.Close();
                        
        }
    }
}
