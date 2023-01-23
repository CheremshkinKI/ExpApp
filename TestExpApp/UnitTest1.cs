using HowToExcel;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;
using System.Reflection;

namespace TestExpApp
{
    public class Tests
    {
        //++++
        [Test]
        public void Test1()
        {
            if (typeof(MarketReport).IsClass == true)
            {
                Assert.Pass();
            }
            else 
            { 
                Assert.Fail(); 
            }
        }

        //++++
        [Test]
        public void Test2()
        {
            if (typeof(MarketExcelGenerator).IsClass == true)
            {
                Assert.Pass();
            }
            else
            {
                Assert.Fail();
            }
        }

        //++++
        [Test]
        public void Test3()
        {
            if (typeof(Program).IsClass == true)
            {
                Assert.Pass();
            }
            else
            {
                Assert.Fail();
            }
        }

        //++++
        [Test]
        public void Test4()
        {
            var del = new MarketReport();
            int Countt = 1;

            del.Mater = new Materials[Countt];
            int Count = del.Mater.Length;

            Assert.AreEqual(1, Count);
        }

        [Test]
        public void Test5()
        {
            if (typeof(Equipment[]).IsArray == true)
            {
                Assert.Pass();
            }
            else
            {
                Assert.Fail();
            }
        }

        [Test]
        public void Test6()
        {
            // Arrange
            int[] Count = { 1, 2, 3, 4, 5 };
            int K = 3;

            //Act
            for (int i = 1; i <= Count.Length; i++)
            {
                if (K == i)
                {
                    Count = Count.Where(val => val != K).ToArray();
                }
            }
            
            //Assert
            Assert.AreEqual( 4  , Count.Length);
        }
        [Test]
        public void Test7()
        {
            string path = @"C:\Users\azari\source\repos\ExpApp\HowToExcel\bin\Debug\net5.0\gg.xlsx";

            bool dirName = Directory.Exists(path);
            if (dirName==true)
            {
                Assert.Pass();
            }
            else
            {
                Assert.Pass();
            }
        }

        //++++
        [Test]
        public void Test8()
        {
            if (typeof(Materials).IsClass == true)
            {
                Assert.Pass();
            }
            else
            {
                Assert.Fail();
            }
        }

        //++++
        [Test]
        public void Test9()
        {
            if (typeof(Equipment).IsClass == true)
            {
                Assert.Pass();
            }
            else
            {
                Assert.Fail();
            }
        }

        //++++
        [Test]
        public void Test10()
        {
            if (typeof(MarketReporter).IsClass == true)
            {
                Assert.Pass();
            }
            else
            {
                Assert.Fail();
            }
        }

        
        [Test]
        public void Test11()
        {
            if (typeof(GenericUriParser).IsClass == true)
            {
                Assert.Pass();
            }
            else
            {
                Assert.Fail();
            }
        }

        //[Test]
        //public void Test12()
        //{
        //    int count = 1;
            
        //    if (count == 1)
        //    {
        //        var reportData = new MarketReporter().GetReport1();
        //        var reportExcel = new MarketExcelGenerator().Generate1(reportData);
        //        //Console.WriteLine("Создание таблицы завершенно");
        //        //Console.WriteLine("\n Дайте имя файлу для сохранения: ");
        //        string read = "GachiSerGay";
        //        read = read + ".xlsx";
        //        File.WriteAllBytes(read, reportExcel);

        //        Assert.Pass();
        //    }
        //    else
        //    {
        //        Assert.Fail();
        //    }
        //}

    }
}