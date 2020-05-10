using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using DA;
using DM;
using System.Data;

namespace Yahoo
{
    class  Program
    {
        private static string _cs = "Server=Nixon,1466;Database=Barchart;User Id=sa;Password=@a88word";
        public YahooDa yda = new YahooDa();
        public DayHighDa dhda = new DayHighDa();
        static void Main(string[] args)
        {
            Program p = new Program();
            p.run();
            
           
        }

        public void run()
        {
            dhda.DownloadHistory("FB");
            ////var symbols = yda.GetDistinctSymbolList();
            ////using (SqlCommand comm = new SqlCommand())
            ////{
            ////    using (SqlConnection conn = new SqlConnection(_cs))
            ////    {
            ////        comm.CommandText = "Select Distinct Symbol From ZacksRank";
            ////        comm.Connection = conn;

            ////        conn.Open();
            ////        var dr = comm.ExecuteReader(CommandBehavior.CloseConnection);
            ////        while (dr.Read())
            ////        {
            ////            dhda.DownloadHistory(dr["Symbol"].ToString());
            ////        }
            ////    }
            ////}
        }

        ////public static void DownloadHistory(string symbol)
        ////{
        ////    Console.WriteLine(symbol); 
        ////    try
        ////    {
        ////        SettingsDa sda = new SettingsDa();
        ////        string p1 = sda.GetSetting("Period1");
        ////        string p2 = sda.GetSetting("Period2");
        ////        var wr = WebRequest.Create(
        ////            $"http://query1.finance.yahoo.com/v7/finance/download/{symbol}?period1={p1}&period2={p2}&interval=1d&events=history");
        ////        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        ////        var resp = wr.GetResponse();

        ////        var sr = new StreamReader(resp.GetResponseStream());



        ////        System.IO.File.AppendAllText($"C:\\Users\\dlolf\\Downloads\\{symbol}.csv", sr.ReadToEnd());
        ////        UploadData($"C:\\Users\\dlolf\\Downloads\\{symbol}.csv", symbol);
        ////        //Microsoft.VisualBasic.FileSystem.Rename($"C:\\Users\\dlolf\\Downloads\\{symbol}.csvx", $"C:\\Users\\dlolf\\Downloads\\{symbol}.csv");

        ////    }
        ////    catch (Exception ex)
        ////    {
        ////        Console.WriteLine(ex.Message);
        ////    }
        ////}

        ////public static void UploadData(string path, string symbol)
        ////{
        ////    string sqlCommand = string.Empty;

        ////    //string[] files = Directory.GetFiles(@"C:\Users\dlolf\Downloads", "*.csv");
        ////    //foreach (var file in files)
        ////    //{
        ////    try
        ////    {


        ////        var lines = System.IO.File.ReadAllLines(path);

        ////        bool header = true;
        ////        foreach (var line in lines)
        ////        {
        ////            try
        ////            {
        ////                if (!header)
        ////                {
        ////                    var fields = line.Split(",".ToCharArray());
        ////                    Record rec = new Record();
        ////                    rec.Date = fields[0];
        ////                    rec.Open = fields[1];
        ////                    rec.High = fields[2];
        ////                    rec.Low = fields[3];
        ////                    rec.Close = fields[4];
        ////                    rec.Adjclose = fields[5];
        ////                    rec.Volume = fields[6];

        ////                    if (yda.SymbolDateExist(symbol, rec.Date))
        ////                    {
        ////                        sqlCommand =
        ////                            $"Update [ZacksRank] Set DayHigh = '{rec.High}', [Open] = '{rec.Open}', [Close] = '{rec.Close}', [DayLow] = '{rec.Close}', Volume = '{rec.Volume}' Where Symbol = '{symbol}' And [Date] = '{rec.Date}'";
        ////                    }
        ////                    else
        ////                    {
        ////                        sqlCommand =
        ////                            $"Insert Into [ZacksRank] (Symbol, [Date], DayHigh, [Open], [Close], DayLow, Volume) Values('{symbol}','{rec.Date}','{rec.High}','{rec.Open}','{rec.Close}', '{rec.Low}', '{rec.Volume}')";
        ////                    }

        ////                    yda.ExecuteSqlCommand(sqlCommand);
        ////                }

        ////                header = false;
        ////            }
        ////            catch (Exception ex)
        ////            {
        ////                Console.WriteLine(ex.Message);
        ////            }
        ////        }

        ////        System.IO.File.Delete(path);
        ////    }
        ////    catch (Exception ex)
        ////    {
        ////        Console.WriteLine(ex.Message);
        ////    }
        ////}
    }
}