using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace HowToExcel
{
    public class MarketReporter
    {
        public MarketReport GetReport()
        {
            Console.Write("Введите количество строк: ");
            int count = Convert.ToInt32(Console.ReadLine());

            MarketReport mark = new MarketReport();
            mark.Mater = new Materials [count];
            for (int i = 0; i < count; i++)
            {
                Console.WriteLine("Номер строки "+ (i+1));
                mark.Mater[i] = new Materials();
                Console.Write("Введите вид: ");
                mark.Mater[i].Type = Console.ReadLine();
                Console.Write("Введите марку: ");
                mark.Mater[i].Brand = Console.ReadLine();
                Console.Write("Введите цвет: ");
                mark.Mater[i].Colour = Console.ReadLine();
                Console.Write("Введите размер: ");
                mark.Mater[i].Size = Console.ReadLine();
                Console.Write("Введите Кол-во в одной пачке: ");
                mark.Mater[i].Count = Convert.ToInt32(Console.ReadLine());
                Console.Write("Введите сколько в наличии: ");
                mark.Mater[i].Available = Convert.ToInt32(Console.ReadLine());
                Console.Write("Введите цену: ");
                mark.Mater[i].Price = Convert.ToInt32(Console.ReadLine());
                Console.Write("Введите продавца: ");
                mark.Mater[i].Seller = Console.ReadLine();
                Console.Write("Введите ссылка: ");
                mark.Mater[i].Link = Console.ReadLine();
            }
            return mark;
        }

        public MarketReport GetReport1()
        {
            Console.Write("Введите количество строк: ");
            int count = Convert.ToInt32(Console.ReadLine());

            MarketReport mark = new MarketReport();
            mark.Equip = new Equipment[count];
            for (int i = 0; i < count; i++)
            {
                Console.WriteLine("Номер строки " + (i + 1));
                mark.Equip[i] = new Equipment();
                Console.Write("Введите вид: ");
                mark.Equip[i].Name = Console.ReadLine();
                Console.Write("Введите марку: ");
                mark.Equip[i].Price = Console.ReadLine();
                Console.Write("Введите цвет: ");
                mark.Equip[i].Unit = Console.ReadLine();
                Console.Write("Введите Кол-во в одной пачке: ");
                mark.Equip[i].Resource = Convert.ToInt32(Console.ReadLine());
            }
            return mark;
        }

        public MarketReport GetReportEmpty1()
        {
            //Console.Write("Введите количество строк: ");
            int count = 1; //Convert.ToInt32(Console.ReadLine());
            string NameEmpty = " ";
            string PriceEmpty = " ";
            string UnitEmpty = " ";
            int ResourceEmpty = 0;

            MarketReport mark = new MarketReport();
            mark.Equip = new Equipment[count];
            for (int i = 0; i < count; i++)
            {
                //Console.WriteLine("Номер строки " + (i + 1));
                //mark.Equip[i] = new Equipment();
                //Console.Write("Введите вид: ");
                mark.Equip[i].Name = NameEmpty; //Console.ReadLine();
                //Console.Write("Введите марку: ");
                mark.Equip[i].Price = PriceEmpty; //Console.ReadLine();
                //Console.Write("Введите цвет: ");
                mark.Equip[i].Unit = UnitEmpty; //Console.ReadLine();
                //Console.Write("Введите Кол-во в одной пачке: ");
                mark.Equip[i].Resource = ResourceEmpty; //Convert.ToInt32(Console.ReadLine());
            }
            return mark;
        }

    }


    public class Program
    {
        
        private double balance;
        public Program()
        { 
        }
        public Program(double balance)
        {
            this.balance = balance;
        }
        public double Balance
        { 
            get { return balance; }
        }

        public void Add(double amount) 
        {
            if (amount < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(amount));
            }

            balance += amount;

        }

        static void Main(string[] args)
        {
            bool b = true;
            do
            {
                int table;
                int str;
                int menu;
                int korzina;
                /*
                 1.Открыть таблицу 
                    1) Список файлов
                 2.Создать таблицу
                    2) Какую таблицу 
                        а)Материалов = проверка(Только одна таблица одно типа материалов)
                        б)Оборудования = проверка(только одно)
                 3.Удалить таблицу 
                    1) Список таблиц
                 4.Создать корзину
                 */
                string location = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string dirName = Path.GetDirectoryName(location);
                Console.WriteLine("\n Меню \n 1-открыть таблицу \n 2-Создать таблицу \n 3-Удалить таблицу \n 4-Создать корзину \n 5-Выход");
                menu = Convert.ToInt32(Console.ReadLine());
                switch (menu)
                {
                    //Открытие таблицы в консоли
                    case 1:
                        int y = 1;
                        int County;

                        if (Directory.Exists(dirName))
                        {
                            Console.WriteLine("Файлы:");
                            string[] files = Directory.GetFiles(dirName, "*.xlsx");
                            foreach (string s in files)
                            {
                                Console.WriteLine(y + " " + s);
                                y++;
                            }

                            Console.Write("Укажите номер файла, который хотите вывести: ");
                            County = Convert.ToInt32(Console.ReadLine()) - 1;

                            for (int i = 1; i <= files.Length; i++)
                            {
                                if (County == i)
                                {

                                    if (File.Exists(files[County]))
                                    {
                                        Application excelApp = new Application();

                                        Workbook excelBook = excelApp.Workbooks.Open(files[County]);
                                        _Worksheet excelSheet = excelBook.Sheets[1];
                                        Range excelRange = excelSheet.UsedRange;

                                        int rows = excelRange.Rows.Count;
                                        int cols = excelRange.Columns.Count;

                                        for (int k = 1; k <= rows; k++)
                                        {
                                            //create new line
                                            Console.Write("\r\n");
                                            for (int j = 1; j <= cols; j++)
                                            {

                                                //write the console
                                                if (excelRange.Cells[k, j] != null && excelRange.Cells[k, j].Value2 != null)
                                                    Console.Write(excelRange.Cells[k, j].Value2.ToString() + "\t");
                                            }
                                        }
                                        excelApp.Quit();
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                                        Console.ReadLine();
                                    }
                                    else
                                    {
                                        Console.WriteLine("Файл не существует");
                                    }
                                }

                            }
                        }
                        break;
                    //создание таблицы
                    case 2:
                        Console.WriteLine("\n Какую таблицу создать? \n 1-для материалов \n 2-для оборудования");
                        table = Convert.ToInt32(Console.ReadLine());
                        switch (table)
                        {
                            case 1:
                                Console.WriteLine("Создать строку? \n 1-да \n 2-нет");
                                str = Convert.ToInt32(Console.ReadLine());
                                switch (str)
                                {
                                    case 1:
                                        var reportData = new MarketReporter()
                                                .GetReport();
                                        var reportExcel = new MarketExcelGenerator()
                                                .Generate(reportData);
                                        Console.WriteLine("Создание таблицы завершенно");
                                        //сохранение
                                        Console.WriteLine("\n Дайте имя файлу для сохранения: ");
                                        string read = Console.ReadLine();
                                        read = read + ".xlsx";
                                        File.WriteAllBytes(read, reportExcel);
                                        break;

                                    case 2:

                                        break;
                                }
                                break;

                            case 2:
                                Console.WriteLine("Создать строку? \n 1-да \n 2-нет");
                                str = Convert.ToInt32(Console.ReadLine());
                                switch (str)
                                {
                                    case 1:
                                        var reportData = new MarketReporter()
                                                    .GetReport1();
                                        var reportExcel = new MarketExcelGenerator()
                                                    .Generate1(reportData);
                                        Console.WriteLine("Создание таблицы завершенно");
                                        //сохранение
                                        Console.WriteLine("\n Дайте имя файлу для сохранения: ");
                                        string read = Console.ReadLine();
                                        read = read + ".xlsx";
                                        File.WriteAllBytes(read, reportExcel);
                                        break;

                                    case 2:
                                        var reportData1 = new MarketReporter()
                                                .GetReportEmpty1();
                                        var reportExcel1 = new MarketExcelGenerator()
                                                .GenerateEmpty1(reportData1);
                                        Console.WriteLine("Создание таблицы завершенно");
                                        //сохранение
                                        Console.WriteLine("\n Дайте имя файлу для сохранения: ");
                                        string read1 = Console.ReadLine();
                                        read1 = read1 + ".xlsx";
                                        File.WriteAllBytes(read1, reportExcel1);
                                        break;
                                }
                                break;
                        }
                        break;
                    //Удаление таблицы
                    case 3:
                        
                        int q = 1;
                        int Countt;

                        if (Directory.Exists(dirName))
                        {
                            Console.WriteLine("Файлы:");
                            string[] files = Directory.GetFiles(dirName, "*.xlsx");
                            foreach (string s in files)
                            {
                                Console.WriteLine(q + " " + s);
                                q++;
                            }
                            Console.Write("Укажите номер файла, который хотите удалить: ");
                            Countt = Convert.ToInt32(Console.ReadLine()) - 1;

                            for (int i = 1; i <= files.Length; i++)
                            {
                                if (Countt == i)
                                {

                                    if (File.Exists(files[Countt]))
                                    {
                                        File.Delete(files[Countt]);
                                        Console.WriteLine("Файл удален");
                                    }
                                    else
                                    {
                                        Console.WriteLine("Файл не существует");
                                    }
                                }

                            }
                        }
                        break;
                    //Создание корзины
                    case 4:
                        int Countty;

                        //var reportExcelNewBox = new MarketExcelGenerator()
                        //                        .GenerateNew();
                        //Console.WriteLine("\n Дайте имя файлу для сохранения: ");
                        //string readNewBox = Console.ReadLine();
                        //readNewBox = readNewBox + ".xlsx";
                        //File.WriteAllBytes(readNewBox, reportExcelNewBox);

                    //Console.WriteLine("Введите кол-во файлов, которое хотите ввести: ");
                    //int CountOfFile = Convert.ToInt32(Console.ReadLine());

                        Label:
                            Console.WriteLine("Добавить файл в корзину? \n 1-да \n 2-нет");
                            korzina = Convert.ToInt32(Console.ReadLine());

                            switch (korzina)
                            {
                                case 1:
                                    if (Directory.Exists(dirName))
                                    {
                                        Console.WriteLine("Файлы:");
                                        string[] files = Directory.GetFiles(dirName, "*.xlsx");
                                        int t = 1;
                                        foreach (string s in files)
                                        {
                                            Console.WriteLine(t + " " + s);
                                            t++;
                                        }

                                        Console.Write("Укажите номер файла, который хотите добавить: ");
                                        Countty = Convert.ToInt32(Console.ReadLine()) - 1;

                                        for (int i = 1; i <= files.Length; i++)
                                        {
                                            if (Countty == i)
                                            { 

                                            
                                            if (File.Exists(files[Countty]))
                                            {
                                                Application excelApp = new Application();

                                                Workbook excelBook = excelApp.Workbooks.Open(files[Countty]);
                                                _Worksheet excelSheet = excelBook.Sheets[1];
                                                Range excelRange = excelSheet.UsedRange;
                                                                                               
                                                List<string> termsList = new List<string>();

                                                int rows = excelRange.Rows.Count;
                                                int cols = excelRange.Columns.Count;

                                                for (int k = 1; k <= rows; k++)
                                                {
                                                    //create new line
                                                    Console.Write("\r\n");
                                                    for (int j = 1; j <= cols; j++)
                                                    {
                                                        //write the console
                                                        if (excelRange.Cells[k, j] != null && excelRange.Cells[k, j].Value2 != null)
                                                        {   //Console.Write(excelRange.Cells[k, j].Value2.ToString() + "\t");

                                                            //string[] gtr = excelRange.Columns.Value2;
                                                            //int x = 1;
                                                            //foreach (string s in files)
                                                            //{
                                                            //    Console.WriteLine(x + " " + s);
                                                            //    x++;
                                                            //}


                                                            string temp = excelRange.Cells[k, j].Value2.ToString();
                                                            termsList.Add(temp);

                                                            string[] terms = termsList.ToArray();
                                                            //Console.WriteLine(terms.ToString());
                                                        }
                                                    }
                                                }
                                                Console.Write("\n");
                                                int v = 0;
                                                int m = 1;
                                                foreach (string e in termsList)
                                                {
                                                    if (v % 4 == 0)
                                                    {
                                                       
                                                        Console.Write(m + " " + e + "\t");
                                                        m++;
                                                    }
                                                    else 
                                                    {
                                                       
                                                        Console.Write( e + "\t");
                                                    }
                                                    v++;
                                                    
                                                }


                                                Console.Write("\n Укажите номер строки, который хотите добавить: ");
                                                int number = Convert.ToInt32(Console.ReadLine());

                                                int p = 1;
                                                foreach (string s in files)
                                                {
                                                    Console.WriteLine(p + " " + s);
                                                    t++;
                                                }


                                                excelApp.Quit();
                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                                                Console.ReadLine();
                                            }
                                            

                                        }

                                        }

                                    }
                                    goto Label;

                                case 2:
                                    break;
                            }
                            break;

                    case 5:
                        b = false;
                        break;
                }

            }
            while(b);
            }
        
    }
}
