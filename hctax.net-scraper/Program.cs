﻿using CsvHelper;
using CsvHelper.Configuration.Attributes;
using HtmlAgilityPack;
using OfficeOpenXml;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;

namespace hctax.net_scraper
{
    public class DataModel
    {
        [Index(0)]
        public string FirstName { get; set; }
        [Index(1)]
        public string LastName { get; set; }
        [Index(2)]
        public string Address { get; set; }
        [Index(3)]
        public string City { get; set; }
        [Index(4)]
        public string State { get; set; }
        [Index(5)]
        public string Zip { get; set; }
        [Ignore]
        public string Amount { get; set; }
    }

    // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse);
    public class Record
    {
        public string Account { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public object Status { get; set; }
        public object Reason { get; set; }
    }

    public class RootInit
    {
        public string Result { get; set; }
        public List<Record> Records { get; set; } = new List<Record>();
        public int TotalRecordCount { get; set; }
    }
    public class ProxyServer
    {
        public string proxy { get; set; }
        public string ip { get; set; }
        public string port { get; set; }
        public string connectionType { get; set; }
        public string asn { get; set; }
        public string isp { get; set; }
        public string type { get; set; }
        public int lastChecked { get; set; }
        public bool get { get; set; }
        public bool post { get; set; }
        public bool cookies { get; set; }
        public bool referer { get; set; }
        public bool userAgent { get; set; }
        public string city { get; set; }
        public string state { get; set; }
        public string country { get; set; }
        public string randomUserAgent { get; set; }
        public int requestsRemaining { get; set; }
    }


    internal class Program
    {
        static void Main(string[] args)
        {
            List<DataModel> entries = ReadInputFile();
            int num = 10;
            int pages = (entries.Count + num - 1) / num;
            List<Task> tasks = new List<Task>();
            for (int count = 1; count <= pages; ++count)
            {
                int index = count - 1;
                var data = entries.Skip(index * num).Take(num).ToList();

                Task newTask = Task.Factory.StartNew(() => { ProcessRecord(data).Wait(); });
                tasks.Add(newTask);

                if (count % 20 == 0 || count == pages)
                {
                    // Thread.Sleep(2000);
                    foreach (Task task in tasks)
                    {
                        while (!task.IsCompleted)
                        { }
                    }
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage excel = new ExcelPackage())
                {
                    excel.Workbook.Worksheets.Add("Products").Cells[1, 1].LoadFromCollection(entries, true);
                    excel.SaveAs(new FileInfo("results.xlsx"));
                }
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Products").Cells[1, 1].LoadFromCollection(entries, true);
                excel.SaveAs(new FileInfo("results.xlsx"));
            }
            Console.WriteLine("Operation Completed");
            Console.ReadKey();
        }

        public static WebProxy GetWebProxy()
        {
            WebProxy proxy;
            try
            {
                proxy = new WebProxy("zproxy.lum-superproxy.io", 22225)
                {
                    Credentials = new NetworkCredential("lum-customer-hl_84bcc60a-zone-data_center-country-us", "")
                };
            }
            catch (Exception ex)
            {
                throw new Exception("Error:" + ex.Message);
            }
            return proxy;
        }

        private static async Task ProcessRecord(List<DataModel> model)
        {
            var proxy = GetWebProxy();
            foreach (var data in model)
            {
                string messages = "";
                try
                {
                    messages += "Processed address: " + data.Address + "\n";
                    RestClientOptions options = new RestClientOptions("https://www.hctax.net/Property/Actions/AccountsList?jtStartIndex=0&jtPageSize=20&jtSorting=Name%20ASC")
                    {
                        Proxy = proxy
                    };
                    var client = new RestClient(options);
                    var request = new RestRequest();
                    request.Method = Method.Post;
                    request.AlwaysMultipartFormData = true;
                    request.AddParameter("colSearch", "address");
                    request.AddParameter("searchText", data.Address);
                    var response = await client.ExecuteAsync<RootInit>(request);
                    if (response != null && response.Data != null)
                    {
                        if (response.Data.TotalRecordCount == 0)
                        {
                            Console.WriteLine("No result found");
                        }
                        else if (response.Data.TotalRecordCount == 1)
                        {
                            string id = await GetEncryptedId(response.Data.Records.FirstOrDefault().Account, proxy);
                            id = id.Replace("\r\n", "");
                            if (!string.IsNullOrEmpty(id))
                            {
                                HtmlWeb web = new HtmlWeb();
                                HtmlDocument doc = web.Load($"https://www.hctax.net/Property/TaxStatement?account={id}", "GET", proxy, new NetworkCredential("lum-customer-hl_84bcc60a-zone-data_center-country-us", ""));

                                var node = doc.DocumentNode.SelectNodes("//tr").FirstOrDefault(x => x.InnerText.Contains("Total Amount Due"));
                                if (node != null)
                                {
                                    var parts = node.InnerText.Split(new String[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                                    data.Amount = parts[1].Trim();
                                    messages += "Total Due: " + data.Amount + "\n";
                                }
                            }
                            else
                                messages += "We cant get the details because we are unable to get the encrypted id" + "\n";
                        }
                        else
                        {
                            messages += "Found multiple records: " + response.Data.TotalRecordCount + "\n";
                            foreach (var record in response.Data.Records)
                            {
                                if (record.Name.ToLower().Contains(data.LastName.ToLower()))
                                {
                                    string id = await GetEncryptedId(record.Account, proxy);
                                    id = id.Replace("\r\n", "");
                                    if (!string.IsNullOrEmpty(id))
                                    {
                                        HtmlWeb web = new HtmlWeb();
                                        HtmlDocument doc = web.Load($"https://www.hctax.net/Property/TaxStatement?account={id}", "GET", proxy, new NetworkCredential("lum-customer-hl_84bcc60a-zone-data_center-country-us", ""));

                                        var node = doc.DocumentNode.SelectNodes("//tr").FirstOrDefault(x => x.InnerText.Contains("Total Amount Due"));
                                        if (node == null)
                                        {
                                            node = doc.DocumentNode.SelectNodes("//tr").FirstOrDefault(x => x.InnerText.Contains("Total Due >>>"));
                                        }
                                        if (node != null)
                                        {
                                            var parts = node.InnerText.Split(new String[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Where(x => x != "                            ").ToList();
                                            data.Amount = parts[1].Trim();
                                            messages += "Total Due: " + data.Amount + "\n";
                                        }
                                    }
                                    else
                                        messages += "We cant get the details because we are unable to get the encrypted id" + "\n";
                                    break;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    messages += "We are unable to process address. Reason: " + ex.Message + "\n";
                }
                Console.WriteLine(messages);
            }
        }

        private static async Task<string> GetEncryptedId(string accountNumber, WebProxy proxy)
        {
            RestClientOptions options = new RestClientOptions($"https://www.hctax.net/Property/AccountEncrypt?account={accountNumber}")
            {
                Proxy = proxy
            };
            var client = new RestClient(options);
            var request = new RestRequest();
            request.Method = Method.Get;
            var response = await client.ExecuteAsync(request);
            return response.Content;
        }

        private static List<DataModel> ReadInputFile()
        {
            using (var reader = new StreamReader("file.csv"))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                return csv.GetRecords<DataModel>().ToList();
            }
        }
    }
}
