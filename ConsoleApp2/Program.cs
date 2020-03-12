﻿using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Net;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            ImportBarchart();

            using (SqlCommand comm = new SqlCommand())
            {
                using (SqlConnection conn =
                    new SqlConnection("Server=Localhost,1466;Database=Barchart;User Id=sa;Password=@a88word"))
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

        static void UpdateZacksRank(string Symbol)
        {
            if (!DoesZacksRankExistForToday(Symbol))
            {
                try
                {
                    var wr = WebRequest.Create("https://www.zacks.com/stock/quote/" + Symbol);
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    var resp = wr.GetResponse();

                    var sr = new StreamReader(resp.GetResponseStream());

                    string text = sr.ReadToEnd();
                    int Rank = 6;
                    if (text.Contains("1-Strong Buy")) Rank = 1;
                    if (text.Contains("2-Buy")) Rank = 2;
                    if (text.Contains("3-Hold")) Rank = 3;
                    if (text.Contains("4-Sell")) Rank = 4;
                    if (text.Contains("5-Strong Sell")) Rank = 5;

                    char Momentum = 'F';
                    if (text.Contains("<span class=\"composite_val\">A</span>&nbsp;Momentum")) Momentum = 'A';
                    if (text.Contains("<span class=\"composite_val\">B</span>&nbsp;Momentum")) Momentum = 'B';
                    if (text.Contains("<span class=\"composite_val\">C</span>&nbsp;Momentum")) Momentum = 'C';
                    if (text.Contains("<span class=\"composite_val\">D</span>&nbsp;Momentum")) Momentum = 'D';

                    InsertZacksRating(Symbol, Rank, Momentum);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            //File.AppendAllText(@"C:\Temp\" + Symbol + ".htm", sr.ReadToEnd());
        }

        static bool DoesZacksRankExistForToday(string Symbol)
        {
            using (SqlCommand comm = new SqlCommand())
            {
                using (SqlConnection conn = new SqlConnection("Server=Localhost,1466;Database=Barchart;User Id=sa;Password=@a88word"))
                {
                    comm.CommandText =
                        $"Select Count(1) From ZacksRank Where Symbol = '{Symbol}' And [Date] >= '{DateTime.Now.ToString("yyyy-MM-dd")}'";
                    comm.Connection = conn;
                    conn.Open();

                    return Convert.ToBoolean(comm.ExecuteScalar());
                }
            }
        }

        static void InsertZacksRating(string Symbol, int Rank, char Momentum)
        {
            try
            {
                using (SqlCommand comm = new SqlCommand())
                {
                    using (SqlConnection conn = new SqlConnection("Server=Localhost,1466;Database=Barchart;User Id=sa;Password=@a88word"))
                    {
                        comm.CommandText =
                            $"Insert Into ZacksRank Values('{Symbol}', '{Rank}', '{DateTime.Now.ToString("yyyy-MM-dd")}', '{Momentum}')";
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

        static void ImportBarchart()
        {
            var files = Directory.GetFiles(@"C:\Users\dlolf\Downloads", "*.csv");
            foreach (var file in files)
            {
                Console.WriteLine(Environment.NewLine + file);

                var records = System.IO.File.ReadAllLines(file);
                bool first = true;
                foreach (var record in records)
                {
                    var fields = record.Split(',');

                    if (first)
                    {
                        foreach (var field in fields)
                        {
                            first = false;
                        }
                    }
                    else
                    {
                        if (LookUpSymbol(fields[0]))
                        {
                            Console.WriteLine($"Symbol Already Exists : {fields[0]}");
                        }
                        else if(fields.Length < 3)
                        {
                            Console.WriteLine($"Invalid record : {fields[0]}");
                        }
                        else
                        {
                            DateTime result;
                            if (fields.Length > 9)
                            {
                                if (!DateTime.TryParse(fields[9], out result))
                                {
                                    fields[9] = DateTime.Now.ToString();
                                }
                            }

                            // Build Insert Statement
                            string stmt = "Insert Into Top100 (Symbol, Name, Date) Values(";
                            int i = 0;
                            foreach (var field in fields)
                            {
                                i++;
                                if (i > 2)
                                    break;

                                stmt += $"'{field.Replace("\"", "").Replace(",", "").Replace("%", "").Replace("\'", "").Replace("unch", "0")}',";
                            }
                            //stmt = stmt.Substring(0, stmt.Length - 1);
                            stmt += ("'" + System.DateTime.Now.ToShortDateString() +  "')");

                            Console.WriteLine(stmt);

                            InsertTop100(stmt);
                        }

                    }
                }
            }
        }

        static Boolean LookUpSymbol(string symbol)
        {
            try
            {
                using (SqlCommand comm = new SqlCommand())
                {
                    using (SqlConnection conn =
                        new SqlConnection("Server=Localhost,1466;Database=Barchart;User Id=sa;Password=@a88word"))
                    {
                        comm.CommandText = $"Select Count(1) from Top100 Where Symbol = '{symbol}'";
                        comm.Connection = conn;
                        conn.Open();

                        return Convert.ToBoolean(comm.ExecuteScalar());
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);

            }
            return false;
        }

        static void InsertTop100(string stmt)
        {
            try
            {
                using (SqlCommand comm = new SqlCommand())
                {
                    using (SqlConnection conn =
                        new SqlConnection("Server=Localhost,1466;Database=Barchart;User Id=sa;Password=@a88word"))
                    {
                        comm.CommandText = stmt;
                        comm.Connection = conn;
                        conn.Open();

                        comm.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);

            }

        }
    }
}