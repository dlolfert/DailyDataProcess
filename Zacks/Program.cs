using System;
using System.CodeDom;
using System.Data.SqlClient;
using System.IO;
using System.Net;

namespace Zacks
{
    class Program
    {
        static private string _cs = "Server=Nixon,1466;Database=Barchart;User Id=sa;Password=@a88word";
        
		static void Main(string[] args)
        {
            //ImportBarchart();

            using (SqlCommand comm = new SqlCommand())
            {
                using (SqlConnection conn =
                    new SqlConnection(_cs))
                {
                    comm.CommandText = "Select Symbol From Top100";
                    comm.Connection = conn;
                    conn.Open();

                    var dr = comm.ExecuteReader();
                    while (dr.Read())
                    {
                        UpdateZacksRank(Convert.ToString(dr[0]));
                    }
                }
            }

            //Console.WriteLine("");
            //Console.WriteLine("Press Any Key to contiue!");
            //Console.ReadKey();
        }

        static void UpdateZacksRank(string symbol)
        {
            if (!DoesZacksRankExistForToday(symbol))
            {
                try
                {
                    var wr = WebRequest.Create("https://www.zacks.com/stock/quote/" + symbol);
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    var resp = wr.GetResponse();

                    var sr = new StreamReader(resp.GetResponseStream());

                    string text = sr.ReadToEnd();
                    
                    text = text.Replace((char) '\n', ' ');
                    text = text.Replace('\r', ' ');
                    text = text.Replace("> <", "><");
                    text = text.Replace(">  <", "><");
                    text = text.Replace(">   <", "><");
                    text = text.Replace(">    <", "><");
                    text = text.Replace(">     <", "><");
                    text = text.Replace(">      <", "><");
                    text = text.Replace(">       <", "><");
                    text = text.Replace(">        <", "><");
                    text = text.Replace(">         <", "><");
                    text = text.Replace(">          <", "><");

                    int rank = 6;
                    if (text.Contains("1-Strong Buy")) rank = 1;
                    if (text.Contains("2-Buy")) rank = 2;
                    if (text.Contains("3-Hold")) rank = 3;
                    if (text.Contains("4-Sell")) rank = 4;
                    if (text.Contains("5-Strong Sell")) rank = 5;

                    char momentum = 'F';
                    if (text.Contains("<span class=\"composite_val\">A</span>&nbsp;Momentum ")) momentum = 'A';
                    if (text.Contains("<span class=\"composite_val\">B</span>&nbsp;Momentum")) momentum = 'B';
                    if (text.Contains("<span class=\"composite_val\">C</span>&nbsp;Momentum")) momentum = 'C';
                    if (text.Contains("<span class=\"composite_val\">D</span>&nbsp;Momentum")) momentum = 'D';

                    #region close
                    string close = "0.00";
                    try
                    {
                        if (text.Contains("<div id=\"get_last_price\" class=\"hide\">"))
                        {
                            int index = text.IndexOf("get_last_price\" class=\"hide\">",
                                StringComparison.InvariantCulture);

                            close = text.Substring(index + 29, 100);
                            index = close.IndexOf("</div>", StringComparison.InvariantCulture);
                            close = close.Substring(0, index);
                            decimal dclose;
                            if (!Decimal.TryParse(close, out dclose))
                            {
                                close = "0.00";
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }
                    #endregion

                    #region open

                    string open = "0.00";
                    try
                    {
                        if (text.Contains("Open</dt><dd>"))
                        {
                            int index = text.IndexOf("Open</dt><dd>",
                                StringComparison.InvariantCulture);

                            open = text.Substring(index + 13, 20);
                            index = open.IndexOf("</dd>", StringComparison.InvariantCulture);
                            open = open.Substring(0, index);
                            decimal dclose;
                            if (!Decimal.TryParse(open, out dclose))
                            {
                                open = "0.00";
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }


                    #endregion 

                    #region day high
                    string dayhigh = "0.00";
                    try
                    {
                        if (text.Contains("Day High</dt><dd>"))
                        {
                            int index = text.IndexOf("Day High</dt><dd>",
                                StringComparison.InvariantCulture);

                            dayhigh = text.Substring(index + 17, 20);
                            index = dayhigh.IndexOf("</dd>", StringComparison.InvariantCulture);
                            dayhigh = dayhigh.Substring(0, index);
                            decimal dclose;
                            if (!Decimal.TryParse(dayhigh, out dclose))
                            {
                                dayhigh = "0.00";
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }
                    #endregion

                    #region day low

                    string daylow = "0.00";
                    try
                    {
                        if (text.Contains("Day Low</dt><dd>"))
                        {
                            int index = text.IndexOf("Day Low</dt><dd>",
                                StringComparison.InvariantCulture);

                            daylow = text.Substring(index + 16, 20);
                            index = daylow.IndexOf("</dd>", StringComparison.InvariantCulture);
                            daylow = daylow.Substring(0, index);
                            decimal ddaylow;
                            if (!Decimal.TryParse(daylow, out ddaylow))
                            {
                                daylow = "0.00";
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }
                    #endregion

                    #region volume
                    //volume" class="hide">
                    string volume = "0.00";
                    try
                    {
                        if (text.Contains("volume\" class=\"hide\">"))
                        {
                            int index = text.IndexOf("volume\" class=\"hide\">",
                                StringComparison.InvariantCulture);

                            volume = text.Substring(index + 21, 20);
                            index = volume.IndexOf("</div>", StringComparison.InvariantCulture);
                            volume = volume.Substring(0, index);
                            volume = volume.Replace(",", "");

                            decimal dvolume;
                            if (!Decimal.TryParse(volume, out dvolume))
                            {
                                volume = "0.00";
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }
                    #endregion

                    InsertZacksRating(symbol, rank, momentum, dayhigh, open, close, daylow, volume);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            //File.AppendAllText(@"C:\Temp\" + Symbol + ".htm", sr.ReadToEnd());
        }

        static bool DoesZacksRankExistForToday(string symbol)
        {
            using (SqlCommand comm = new SqlCommand())
            {
                using (SqlConnection conn = new SqlConnection(_cs))
                {
                    comm.CommandText =
                        $"Select Count(1) From ZacksRank Where Symbol = '{symbol}' And [Date] >= '{DateTime.Now.ToString("yyyy-MM-dd")}'";
                    comm.Connection = conn;
                    conn.Open();

                    return Convert.ToBoolean(comm.ExecuteScalar());
                }
            }
        }

        static void InsertZacksRating(string symbol, int rank, char momentum, string dayHigh, string open, string close, string daylow, string volume)
        {
            try
            {
                using (SqlCommand comm = new SqlCommand())
                {
                    using (SqlConnection conn = new SqlConnection(_cs))
                    {
                        comm.CommandText =
                            $"Insert Into ZacksRank Values('{symbol}', '{rank}', '{DateTime.Now.ToString("yyyy-MM-dd")}', '{momentum}', {dayHigh}, '{open}', '{close}', '{daylow}', '{volume}')";
                        comm.Connection = conn;
                        conn.Open();

                        comm.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        
    }
}