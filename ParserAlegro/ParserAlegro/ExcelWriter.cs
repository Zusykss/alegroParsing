using System;
using System.IO;
using OfficeOpenXml;

namespace ParserAlegro
{
    class ExcelWriter
    {
        private readonly string filePath;

        public ExcelWriter(string _filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            filePath = _filePath;
            //if (File.Exists(_filePath))
            //    return;

            ExcelPackage ExcelPkg = new ExcelPackage();
            ExcelWorksheet wsSheet = ExcelPkg.Workbook.Worksheets.Add("List1");
            wsSheet.Protection.IsProtected = false;
            wsSheet.Protection.AllowSelectLockedCells = false;
            ExcelPkg.SaveAs(new FileInfo(_filePath));

            this.filePath = _filePath;
            var fi = new FileInfo(_filePath);
            using (var package = new ExcelPackage(fi))
            {
                package.Workbook.Worksheets.Add("result");
                int indexRows = 1;
                using (var workSheet = package.Workbook.Worksheets["result"])
                {
                    try
                    {
                        var i = workSheet.Dimension.End.Row;
                        indexRows = workSheet.Dimension.End.Row;
                        if (indexRows > 0) indexRows++;
                    }
                    catch { }
                    workSheet.Cells[indexRows, 1].Value = "виробник";
                    workSheet.Cells[indexRows, 2].Value = "oferta";
                    workSheet.Cells[indexRows, 3].Value = "код";
                    workSheet.Cells[indexRows, 4].Value = "ціна";
                    workSheet.Cells[indexRows, 5].Value = "шт";
                    workSheet.Cells[indexRows, 6].Value = "фото";
                    workSheet.Cells[indexRows, 7].Value = "стан";
                    //workSheet.Cells[indexRows, 8].Value = "url";
                    //workSheet.Cells[indexRows, 1].Value = "Название_позиции";
                    //workSheet.Cells[indexRows, 2].Value = "Код_товара";
                    //workSheet.Cells[indexRows, 3].Value = "Ключевые_слова";
                    //workSheet.Cells[indexRows, 4].Value = "Метки";
                    //workSheet.Cells[indexRows, 5].Value = "Описание";
                    //workSheet.Cells[indexRows, 6].Value = "Тип_товара";
                    //workSheet.Cells[indexRows, 7].Value = "Цена";
                    //workSheet.Cells[indexRows, 8].Value = "Валюта";
                    //workSheet.Cells[indexRows, 9].Value = "Скидка";
                    //workSheet.Cells[indexRows, 10].Value = "Единица_измерения";
                    //workSheet.Cells[indexRows, 11].Value = "Минимальный_объем_заказа";
                    //workSheet.Cells[indexRows, 12].Value = "Оптовая_цена";
                    //workSheet.Cells[indexRows, 13].Value = "Минимальный_заказ_опт";
                    //workSheet.Cells[indexRows, 14].Value = "Ссылка_изображения";
                    //workSheet.Cells[indexRows, 15].Value = "Наличие";
                    //workSheet.Cells[indexRows, 16].Value = "Производитель";
                    //workSheet.Cells[indexRows, 17].Value = "Страна_производитель";
                    //workSheet.Cells[indexRows, 18].Value = "Номер_группы";
                    //workSheet.Cells[indexRows, 19].Value = "Идентификатор_группы";
                    //workSheet.Cells[indexRows, 20].Value = "Название_группы";
                    //workSheet.Cells[indexRows, 21].Value = "Адрес_подраздела";
                    //workSheet.Cells[indexRows, 22].Value = "Возможность_поставки";
                    //workSheet.Cells[indexRows, 23].Value = "Срок_поставки";
                    //workSheet.Cells[indexRows, 24].Value = "Способ_упаковки";
                    //workSheet.Cells[indexRows, 25].Value = "Идентификатор_товара";
                    //workSheet.Cells[indexRows, 26].Value = "Уникальный_идентификатор";
                    //workSheet.Cells[indexRows, 27].Value = "Название_Характеристики";
                    //workSheet.Cells[indexRows, 28].Value = "Измерение_Характеристики";
                    //workSheet.Cells[indexRows, 29].Value = "Значение_Характеристики";
                    //workSheet.Cells[indexRows, 30].Value = "Название_Характеристики";
                    //workSheet.Cells[indexRows, 31].Value = "Измерение_Характеристики";
                    //workSheet.Cells[indexRows, 32].Value = "Значение_Характеристики";
                    //workSheet.Cells[indexRows, 33].Value = "";
                    //workSheet.Cells[indexRows, 34].Value = "";
                    //workSheet.Cells[indexRows, 35].Value = "";
                    package.Save();
                }
            }
        }

        public void WriteRow(string title, string price, string numLot, string catalogNum, string url, string photos)
        {
            var fi = new FileInfo(filePath);
            using (var package = new ExcelPackage(fi))
            {
                int indexRows = 1;
                using (var workSheet = package.Workbook.Worksheets["result"])
                {
                    try
                    {
                        var i = workSheet.Dimension.End.Row;
                        indexRows = workSheet.Dimension.End.Row;
                        if (indexRows > 0) indexRows++;
                    }
                    catch { }
                    
                    //workSheet.Cells[indexRows, 1].Value = title;
                    //workSheet.Cells[indexRows, 2].Value = numLot;
                    //workSheet.Cells[indexRows, 4].Value = "allegro";
                    //workSheet.Cells[indexRows, 6].Value = "r";
                    //workSheet.Cells[indexRows, 7].Value = price;
                    //workSheet.Cells[indexRows, 8].Value = "PLN";
                    //workSheet.Cells[indexRows, 10].Value = "шт.";
                    //workSheet.Cells[indexRows, 14].Value = photos.Replace("/s360/", "/s2048/");
                    //workSheet.Cells[indexRows, 15].Value = "1";
                    //workSheet.Cells[indexRows, 25].Value = numLot;
                    //workSheet.Cells[indexRows, 26].Value = numLot;
                    //workSheet.Cells[indexRows, 27].Value = "Номер по каталогу деталей";
                    //workSheet.Cells[indexRows, 29].Value = catalogNum;
                    //workSheet.Cells[indexRows, 30].Value = "Номер по каталогу деталей";
                   // workSheet.Cells[indexRows, 29].Value = catalogNum;

                    // workSheet.Cells[indexRows, 4].Value = catalogNum;

                    package.Save();
                }                
            }
        }
        public void WriteList(List<ObjectAlegro> objectAlegros)
        {
            var fi = new FileInfo(filePath);
            using (var package = new ExcelPackage(fi))
            {
                int indexRows = 1;
                using (var workSheet = package.Workbook.Worksheets["result"])
                {
                    try
                    {
                        var i = workSheet.Dimension.End.Row;
                        indexRows = workSheet.Dimension.End.Row;
                        if (indexRows > 0) indexRows++;
                    }
                    catch { }
                    for (int i = 0; i < objectAlegros.Count; i++)
                    {
                        string photos = "";
                        for (int k = 0; k < objectAlegros[i].Photos.Count; k++)
                        {
                            //Console.WriteLine(obj["medium"].ToString().Replace("\\u002F", "/"));
                            //objectAlegro.Photos.Add(obj["medium"].ToString().Replace("\\u002F", "/"));
                            photos += objectAlegros[i].Photos[k].Replace("\\u002F", "/").Replace("/s360/", "/s2048/");
                            if (k+1 != objectAlegros[i].Photos.Count) photos += ",";
                        }
                        string join = String.Join(',', objectAlegros[i].Photos).Replace("/s360/", "/s2048/");
                        workSheet.Cells[indexRows, 1].Value = objectAlegros[i].Producer;
                        if(string.IsNullOrEmpty(objectAlegros[i].Url)) workSheet.Cells[indexRows, 2].Value = objectAlegros[i].NumberLot;
                        else workSheet.Cells[indexRows, 2].Value = objectAlegros[i].Url.Replace("\\u002F", "/");
                        workSheet.Cells[indexRows, 3].Value = objectAlegros[i].CatalogNumber;
                        workSheet.Cells[indexRows, 4].Value = Convert.ToDouble(objectAlegros[i].Price.Replace(".", ","));
                        workSheet.Cells[indexRows, 5].Value = objectAlegros[i].Quantity;
                        workSheet.Cells[indexRows, 6].Value = photos;
                        workSheet.Cells[indexRows, 7].Value = "used";
                        //workSheet.Cells[indexRows, 10].Value = "шт.";
                        //workSheet.Cells[indexRows, 14].Value = photos; 
                        //workSheet.Cells[indexRows, 15].Value = "1";
                        //workSheet.Cells[indexRows, 25].Value = objectAlegros[i].NumberLot;
                        //workSheet.Cells[indexRows, 26].Value = objectAlegros[i].NumberLot;
                        //workSheet.Cells[indexRows, 27].Value = "Номер по каталогу деталей";
                        //workSheet.Cells[indexRows, 29].Value = objectAlegros[i].CatalogNumber;
                        //workSheet.Cells[indexRows, 30].Value = "Номер по каталогу деталей";
                        //workSheet.Cells[indexRows, 32].Value = objectAlegros[i].CatalogNumber_2;
                        //workSheet.Cells[indexRows, 33].Value = "Номер по каталогу деталей";
                        //workSheet.Cells[indexRows, 35].Value = objectAlegros[i].CatalogNumber_3;

                        indexRows++;
                    }
                    // workSheet.Cells[indexRows, 4].Value = catalogNum;

                    package.Save();
                }
            }
        }
    }
}
