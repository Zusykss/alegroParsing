using ParserAlegro;
using System.Net;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
//using System.Web.Script.Serialization;

//JavaScriptSerializer jsonSerializer = new JavaScriptSerializer();

// See https://aka.ms/new-console-template for more information
ParserAlegro.ParserAlegro parserAlegro = new ParserAlegro.ParserAlegro();
List<string> list = new List<string>();
list.AddRange(File.ReadAllLines("urls.txt", Encoding.UTF8));
List<CatalogUrls> catalogUrls = new List<CatalogUrls>();
bool newCatalog = false;
for (int i = 0; i < list.Count; i++)
{
    if (!list[i].Contains("https"))
    {
        newCatalog = true;
        catalogUrls.Add(new CatalogUrls { NameCatalog = list[i] });
    }
    else
    {
        catalogUrls.Last().CatalogUrs.Add(list[i]);
        newCatalog = false;
    }
}

//list.Add("https://allegro.pl/kategoria/wyposazenie-wnetrza-nawigacje-gps-fabryczne-250553?stan=u%C5%BCywane&price_from=401&price_to=700");
//list.Add("https://allegro.pl/kategoria/wyposazenie-wnetrza-nawigacje-gps-fabryczne-250553?stan=u%C5%BCywane&price_from=201&price_to=400");
//list.Add("https://allegro.pl/kategoria/wyposazenie-wnetrza-nawigacje-gps-fabryczne-250553?stan=u%C5%BCywane&price_from=100&price_to=200");
if (!Directory.Exists("result"))
    Directory.CreateDirectory("result");
string patternCount = "(?<=\"availableCount\":).*?(?=,)";
string patternJson = @"\{\\""aboveTheFoldCount.*?variantsVisible\\"":true}";
for (int ii = 0; ii < catalogUrls.Count; ii++)
{
    list = catalogUrls[ii].CatalogUrs;
    List< ObjectAlegro > listAlegro = new List< ObjectAlegro >();
    //if(!File.Exists("result.xlsx"))
    //    File.Copy("fileXls\\Пример.xlsx", "result.xlsx");

    ExcelWriter excelWriter = new ExcelWriter($"result\\{catalogUrls[ii].NameCatalog}.xlsx");
    List<string> proxies = File.ReadAllLines("proxy.txt").ToList<string>();
    List<string> ignorList = File.ReadAllLines("ignorList.txt").ToList<string>();
    List<string> replaceList = File.ReadAllLines("replaceList.txt").ToList<string>();
    HttpClient httpClient = new HttpClient();
    GoogleSheetsHelper googleSheetsHelper = new GoogleSheetsHelper("1aHcemfVKtd3BUB788ajRgYKP83F3rZ-JclZXSPHgVhQ", "108807396263-hj0n6krt22a1cn25v19frecnv5h5snha.apps.googleusercontent.com", "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");
    // res = await httpClient.GetStringAsync("http://node-pl-3.astroproxy.com:10359/api/changeIP?apiToken=aeb5efa3ded55aa2");
    string urlChangeApi = "";
    for (int i = 0; i < list.Count; i++)
    {
        try
        {
            string proxy = proxies.FirstOrDefault();
            parserAlegro.SerProxy(proxy);
            var listProxy = proxy.Split('@');
            urlChangeApi = listProxy[2];
            var result = await parserAlegro.BrowserLoader(list[i], 0);
            while (result.IndexOf("ERROR BrowserLoader") != -1 || result.IndexOf("Please enable JS and disable any ad blocker") != -1 || result.IndexOf("googlebot\" content=\"noindex, noarchive") != -1)
            {
                var res = await httpClient.GetStringAsync(urlChangeApi);
                Thread.Sleep(1000);
                result = await parserAlegro.BrowserLoader(list[i], 0);
            }
            var count = Regex.Match(result, patternCount).Value;
            var countPages = Convert.ToInt32(count) / 60;
            if (countPages == 0 && Convert.ToInt32(count) > 0) countPages = 1;
            for (int jj = 1; jj <= countPages; jj++)
            {
                //if (jj > 5) break;
                if(jj > 1)
                {
                    try
                    {
                        result = await parserAlegro.BrowserLoader(list[i] + $"&p={jj}", 0);
                        int countRepet = 0;
                        while (result.IndexOf("ERROR BrowserLoader") != -1 || result.IndexOf("Please enable JS and disable any ad blocker") != -1 || result.IndexOf("googlebot\" content=\"noindex, noarchive") != -1) // Please enable JS and disable any ad blocker
                        {
                            var res = await httpClient.GetStringAsync(urlChangeApi);
                            Thread.Sleep(1000);
                            result = await parserAlegro.BrowserLoader(list[i] + $"&p={jj}", 0);
                            Console.WriteLine(countRepet);
                            countRepet++;
                        }
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Contains("Too many requests") || ex.Message.ToLower().Contains("many requests"))
                        {
                            Thread.Sleep(5000);
                            jj--;
                            continue;
                        }
                    }
                }
                var jsonString = Regex.Match(result, @"(?<=\{""listingType"":""base"",""__listing_StoreState"":"").*(?=""}</script>)").Value.Replace("\\\"", "\"");
                jsonString = jsonString.Replace("\\\\\"", " ");
                var jsonObject = new JsonObject();
                try
                {
                    jsonObject = JsonNode.Parse(jsonString).AsObject();
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Console.WriteLine("BadParsing Json write to developer!");
                    //break;
                }
                var array = jsonObject["items"]["elements"].AsArray();
                for (int j = 0; j < array.Count; j++)
                {
                    try
                    {
                        ObjectAlegro objectAlegro = new ObjectAlegro();
                        objectAlegro.Price = array[j]["price"]["normal"]["amount"].ToString();
                        var textInfo = array[j]["deliveryInfo"]["text"].ToString();
                        var pricee = textInfo.Substring(0, textInfo.IndexOf(' '));//.Replace(",", ".");
                        if (double.TryParse(pricee, out double value))
                        {
                            Console.WriteLine($"Price: {value}");
                            objectAlegro.Price = pricee.Replace(",", ".");
                        }
                        objectAlegro.Title = array[j]["title"]["text"].ToString();
                        objectAlegro.NumberLot = array[j]["id"].ToString();
                        string url = array[j]["url"].ToString();
                        if (url.Contains("allegrolokalnie.pl"))
                        {
                            objectAlegro.Url = url;
                        }
                        if (objectAlegro.NumberLot.Length == 0)
                        {
                            if (url.IndexOf("?") != -1)
                            {
                                int indexEnd = url.IndexOf("?");
                                int indexStart = url.LastIndexOf("-", indexEnd);
                                if (indexStart != -1)
                                {
                                    string Id = url.Substring(indexStart, indexEnd - indexStart);
                                    Console.WriteLine(Id);
                                }
                            }
                        }
                        var arrayParam = array[j]["parameters"].AsArray();
                        for (int k = 0; k < arrayParam.Count; k++)
                        {
                            if (arrayParam[k]["name"].ToString() == "Numer katalogowy części")
                            {
                                objectAlegro.CatalogNumber = arrayParam[k]["values"][0].ToString();
                            }
                            else if(arrayParam[k]["name"].ToString() == "Producent części")
                            {
                                objectAlegro.Producer = arrayParam[k]["values"][0].ToString();
                            }
                        }
                        if (!string.IsNullOrEmpty(objectAlegro.Producer))
                        {
                            for (int ii2 = 0; ii2 < replaceList.Count; ii2++)
                            {
                                objectAlegro.Producer = objectAlegro.Producer.Replace(replaceList[ii2], "");
                            }
                        }
                        var arrayPhoto = array[j]["photos"].AsArray();
                        //if (!Directory.Exists($"images\\{objectAlegro.NumberLot}"))
                        //    Directory.CreateDirectory($"images\\{objectAlegro.NumberLot}");
                        string photos = "";
                        for (int k = 0; k < arrayPhoto.Count; k++)
                        {
                            var obj = arrayPhoto[k].AsObject();
                            var str = obj["medium"].ToString();
                            Console.WriteLine(obj["medium"].ToString().Replace("\\u002F", "/"));
                            objectAlegro.Photos.Add(obj["medium"].ToString().Replace("\\u002F", "/") + ".jpg");
                            photos += obj["medium"].ToString().Replace("\\u002F", "/");
                            if (k+1 != arrayPhoto.Count) photos += ",";
                        }
                        objectAlegro.Quantity = array[j]["quantity"].ToString();
                        if(!string.IsNullOrEmpty(objectAlegro.CatalogNumber))
                        {
                            if (!ignorList.Contains(objectAlegro.CatalogNumber))
                            {
                                listAlegro.Add(objectAlegro);
                            }
                        }
                        //break;
                        //googleSheetsHelper.WriteList(listAlegro);
                        //excelWriter.WriteRow(objectAlegro.Title, objectAlegro.Price, objectAlegro.NumberLot, objectAlegro.CatalogNumber, array[j]["url"].ToString().Replace("\\u002F", "/"), photos);
                    }
                    catch(Exception ex)
                    {

                    }
                    //File.AppendAllText("result.csv", objectAlegro.Title + ";" + objectAlegro.Price + ";" + objectAlegro.NumberLot + ";" + objectAlegro.CatalogNumber + "\r\n", Encoding.UTF8);
                }
                //break;  
            }
            Console.WriteLine(listAlegro.Count);
            googleSheetsHelper.WriteList(listAlegro);
            excelWriter.WriteList(listAlegro);
            listAlegro.Clear();
        }
        catch(Exception ex)
        {
            if(ex.Message.Contains("Too many requests") || ex.Message.ToLower().Contains("many requests"))
            {
                Thread.Sleep(30000);

                var res = await httpClient.GetStringAsync(urlChangeApi);

                i--;
                Thread.Sleep(5000);
                continue;
            }
            Console.WriteLine(ex.Message);
        }
    }
}