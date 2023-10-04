using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;

namespace bootcamp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("****************HOŞGELDİNİZ!****************\n");

            InitFiles();

            while (true)
            {
                Console.WriteLine("Lütfen işlem seçin:");
                Console.WriteLine("1-Kasa");
                Console.WriteLine("2-Market");
                Console.WriteLine("3-Patron");
                Console.WriteLine("0-Çıkış");
                string secim = Console.ReadLine();

                switch (secim)
                {
                    case "1":
                        KasaIslemi();
                        break;
                    case "2":
                        MarketIslemi();
                        break;
                    case "3":
                        PatronIslemi();
                        break;
                    case "0":
                        Environment.Exit(0);
                        break;
                    default:
                        Console.WriteLine("Geçersiz seçenek. Lütfen tekrar deneyin.");
                        break;
                }
            }
        }

        static void InitFiles()
        {
            // check if a directory called market prices is exist. if not, create it.
            if (!Directory.Exists("market prices"))
            {
                Directory.CreateDirectory("market prices");
                InitializeMarketPricesFile();

            }
            else
            {
                // if directory exists, check if file exists. if not, create it.
                if (!File.Exists("market prices/market prices.xlsx"))
                {
                    InitializeMarketPricesFile();
                }
            }

            // check if a directory called sales is exist. if not, create it.
            if (!Directory.Exists("sales"))
            {
                Directory.CreateDirectory("sales");
                // create a file called sales.xlsx
                using (File.Create("sales/sales.xlsx")) { }
                // create a sheet called sales
                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("sales");
                worksheet.Cell("A1").Value = "Type";
                worksheet.Cell("B1").Value = "Amount";
                worksheet.Cell("C1").Value = "Date";
                workbook.SaveAs("sales/sales.xlsx");
            }
            else
            {
                // if directory exists, check if file exists. if not, create it.
                if (!File.Exists("sales/sales.xlsx"))
                {
                    using (File.Create("sales/sales.xlsx")) { }
                    // create a sheet called sales
                    var workbook = new XLWorkbook();
                    var worksheet = workbook.Worksheets.Add("sales");
                    worksheet.Cell("A1").Value = "Type";
                    worksheet.Cell("B1").Value = "Amount";
                    worksheet.Cell("C1").Value = "Date";
                    workbook.SaveAs("sales/sales.xlsx");

                }
            }
        }

        static void InitializeMarketPricesFile()
        {
            using (File.Create("market prices/market prices.xlsx")) { }
            // create a sheet called market prices
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("market prices");
            worksheet.Cell("A1").Value = "Product";
            worksheet.Cell("B1").Value = "Price";
            worksheet.Cell("C1").Value = "Number of Sales";
            // add chocolate, water, chips, coke, gum, buscuit as products and their prices
            worksheet.Cell("A2").Value = "Chocolate";
            worksheet.Cell("B2").Value = 5;
            worksheet.Cell("C2").Value = 0;
            worksheet.Cell("A3").Value = "Water";
            worksheet.Cell("B3").Value = 2;
            worksheet.Cell("C3").Value = 0;
            worksheet.Cell("A4").Value = "Chips";
            worksheet.Cell("B4").Value = 3;
            worksheet.Cell("C4").Value = 0;
            worksheet.Cell("A5").Value = "Coke";
            worksheet.Cell("B5").Value = 4;
            worksheet.Cell("C5").Value = 0;
            worksheet.Cell("A6").Value = "Gum";
            worksheet.Cell("B6").Value = 1;
            worksheet.Cell("C6").Value = 0;
            worksheet.Cell("A7").Value = "Buscuit";
            worksheet.Cell("B7").Value = 2;
            worksheet.Cell("C7").Value = 0;
            Console.WriteLine("market prices.xlsx file is updated.");

            workbook.SaveAs("market prices/market prices.xlsx");
        }

        static void KasaIslemi()
        {
            Console.WriteLine("-----------Kasa İşlemi-----------");
            Console.WriteLine("Lütfen satılan benzin miktarını girin(TL):");
            int miktar = int.Parse(Console.ReadLine());

            // Satılan miktarı Excel dosyasına kaydetme işlemi
            KaydetKasaVerisi(miktar);
            Console.Clear();
            Console.WriteLine("Benzin satışı kaydedildi.\n");
        }

        static void MarketIslemi()
        {
            Console.Clear();
            Console.WriteLine("-----------Market İşlemi-----------");
            // Ürünler ve fiyatları Excel dosyasından yüklenir
            Dictionary<string, double> urunler = YukleMarketVerisi();

            if (urunler.Count == 0)
            {
                Console.WriteLine("Ürünler yüklenirken bir hata oluştu.");
                return;
            }

            double toplamFiyat = 0;

            while (true)
            {
                Console.WriteLine("Ürün seçin:");
                foreach (var urun in urunler)
                {
                    Console.WriteLine($"{urun.Key}: {urun.Value}TL");
                }

                Console.WriteLine("Çıkış yapmak için 'q' girin.");
                string secim = Console.ReadLine();

                if (secim.ToLower() == "q")
                {
                    Console.Clear();
                    break;
                }

                if (urunler.ContainsKey(secim))
                {
                    toplamFiyat += urunler[secim];
                    Console.WriteLine($"{secim} sepete eklendi.");
                    Console.Clear();
                    // Ürünlerin satış sayıları Excel dosyasında güncellenir
                    var workbook = new XLWorkbook("market prices/market prices.xlsx");
                    var worksheet = workbook.Worksheet("market prices");
                    int lastRow = worksheet.LastRowUsed().RowNumber();
                    for (int i = 2; i <= lastRow; i++)
                    {
                        if (worksheet.Cell(i, 1).Value.ToString() == secim)
                        {
                            worksheet.Cell(i, 3).Value = worksheet.Cell(i, 3).GetDouble() + 1;
                        }
                    }
                    workbook.SaveAs("market prices/market prices.xlsx");
                    Console.WriteLine(
                        $"Toplam fiyat: {toplamFiyat}TL\n");
                }
                else
                {
                    Console.WriteLine("Geçersiz ürün seçimi.");
                }
            }

            // Yapılan alışverişi Excel dosyasına kaydetme işlemi
            KaydetMarketVerisi(toplamFiyat);

            Console.WriteLine($"Toplam fiyat: {toplamFiyat}TL");
        }

        static void PatronIslemi()
        {
            Console.Clear();
            Console.WriteLine("--------------Patron İşlemi--------------\n");
            Console.WriteLine("1-Satışları Görüntüle");
            Console.WriteLine("2-Market Satış Adetlerini Görüntüle");
            string secim = Console.ReadLine();

            switch (secim)
            {
                case "1":
                    // Satışları görüntüleme işlemi
                    GosterSatisVerisi();
                    break;
                case "2":
                    // Toplam geliri görüntüleme işlemi
                    GosterMarketSatislari();
                    break;
                default:
                    Console.WriteLine("Geçersiz seçenek.");
                    break;
            }
        }

        static void KaydetKasaVerisi(int miktar)
        {
            // save the amount of sold gas to the excel file which is exist in sales directory named sales.xlsx
            var workbook = new XLWorkbook("sales/sales.xlsx");
            var worksheet = workbook.Worksheet("sales");
            int lastRow = worksheet.LastRowUsed().RowNumber();
            worksheet.Cell(lastRow + 1, 1).Value = "Gas";
            worksheet.Cell(lastRow + 1, 2).Value = miktar;
            worksheet.Cell(lastRow + 1, 3).Value = DateTime.Now;
            worksheet.Column(3).AdjustToContents();
            workbook.SaveAs("sales/sales.xlsx");

        }

        static Dictionary<string, double> YukleMarketVerisi()
        {
            Dictionary<string, double> urunler = new Dictionary<string, double>();

            // get the products and their prices from the excel file which is exist in market prices directory named market prices.xlsx
            var workbook = new XLWorkbook("market prices/market prices.xlsx");
            var worksheet = workbook.Worksheet("market prices");
            int lastRow = worksheet.LastRowUsed().RowNumber();
            for (int i = 2; i <= lastRow; i++)
            {
                urunler.Add(worksheet.Cell(i, 1).Value.ToString(), worksheet.Cell(i, 2).GetDouble());
            }

            return urunler;
        }

        static void KaydetMarketVerisi(double toplamFiyat)
        {
            // save the total price of the shopping to the excel file which is exist in sales directory named sales.xlsx
            var workbook = new XLWorkbook("sales/sales.xlsx");
            var worksheet = workbook.Worksheet("sales");
            int lastRow = worksheet.LastRowUsed().RowNumber();
            worksheet.Cell(lastRow + 1, 1).Value = "Shopping";
            worksheet.Cell(lastRow + 1, 2).Value = toplamFiyat;
            worksheet.Cell(lastRow + 1, 3).Value = DateTime.Now;
            worksheet.Column(3).AdjustToContents();
            workbook.SaveAs("sales/sales.xlsx");
        }

        static void GosterSatisVerisi()
        {
            Console.Clear();
            // Show sales from Excel file. show the column names and the values aligned
            var workbook = new XLWorkbook(
                               "sales/sales.xlsx");
            var worksheet = workbook.Worksheet("sales");
            int lastRow = worksheet.LastRowUsed().RowNumber();
            for (int i = 1; i <= lastRow; i++)
            {
                for (int j = 1; j <= 3; j++)
                {
                    Console.Write(worksheet.Cell(i, j).Value.ToString().PadRight(15));
                }
                Console.WriteLine();
            }

            double totalIncome = 0;
            for (int i = 2; i <= lastRow; i++)
            {
                totalIncome += worksheet.Cell(i, 2).GetDouble();
            }
            Console.WriteLine("+______________________________________________________");
            Console.WriteLine($"Toplam gelir: {totalIncome}TL\n\n");
        }

        static void GosterMarketSatislari()
        {
            Console.Clear();
            // Show the market prices from Excel file. show the column names and the values aligned
            var workbook = new XLWorkbook(
                                              "market prices/market prices.xlsx");
            var worksheet = workbook.Worksheet("market prices");
            int lastRow = worksheet.LastRowUsed().RowNumber();
            for (int i = 1; i <= lastRow; i++)
            {
                for (int j = 1; j <= 3; j++)
                {
                    Console.Write(worksheet.Cell(i, j).Value.ToString().PadRight(15));
                }
                Console.WriteLine();
            }
            double totalIncome = 0;
            for (int i = 2; i <= lastRow; i++)
            {
                totalIncome += worksheet.Cell(i, 2).GetDouble();
            }
            Console.WriteLine("+______________________________________________________");
            Console.WriteLine($"Toplam gelir: {totalIncome}TL\n\n");

        }
    }
}
