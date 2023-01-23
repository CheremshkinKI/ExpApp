using System;
namespace HowToExcel
{
    public class MarketReport
    {
        //public Company Company { set; get; }
        public Materials[] Mater { set; get; }
        public Equipment[] Equip { set; get; }
    }
    
    //public class Company
    //{
    //    private int id;
    //    public string Name { set; get; }
    //    public string Address { set; get; }
    //    public string City { set; get; }
    //    public string Country { set; get; }
    //    public string Currency { set; get; }
    //    public string Description { set; get; }
    //    public string Sector { set; get; }
    //}
    
    //public class HistoryItem
    //{
    //    public long Capitalization { set; get; }
    //    public decimal SharePrice { set; get; }
    //    public DateTime Date { set; get; }
    //}

    public class Materials
    { 
        public string Type { set; get; }    //Вид
        public string Brand { set; get; }   //Марка
        public string Colour { set; get; }  //Цвет
        public string Size { set; get; }    //Размер
        public int Count { set; get; }      //Кол-во в одной пачке
        public int Available { set; get; }  // В наличии
        public int Price { set; get; }      //Цена
        public string Seller { set; get; }  //Продавец
        public string Link { set; get; }    //Ссылка
    }

    public class Equipment
    {
        public string Name { set; get; }    //Наименование
        public string Price { set; get; }   //Цена
        public string Unit { set; get; }    //Ед. Измерения
        public int Resource { set; get; }   //Ресурс использования
    }
}