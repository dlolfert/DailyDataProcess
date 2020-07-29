using System;
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
                        if (text.Contains("class=\"alpha\">Open</td>"))
                        {
                            int index = text.IndexOf(">Open</td>",
                                StringComparison.InvariantCulture);

                            open = text.Substring(index + 35, 20);
                            index = open.IndexOf("</td>", StringComparison.InvariantCulture);
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
                        if (text.Contains("class=\"alpha\">Day High</td>"))
                        {
                            int index = text.IndexOf(">Day High</td>",
                                StringComparison.InvariantCulture);

                            dayhigh = text.Substring(index + 39, 20);
                            index = dayhigh.IndexOf("</td>", StringComparison.InvariantCulture);
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

                    InsertZacksRating(symbol, rank, momentum, dayhigh, open, close);
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

        static void InsertZacksRating(string symbol, int rank, char momentum, string dayHigh, string open, string close)
        {
            try
            {
                using (SqlCommand comm = new SqlCommand())
                {
                    using (SqlConnection conn = new SqlConnection(_cs))
                    {
                        comm.CommandText =
                            $"Insert Into ZacksRank Values('{symbol}', '{rank}', '{DateTime.Now.ToString("yyyy-MM-dd")}', '{momentum}', {dayHigh}, '{open}', '{close}', null, null)";
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