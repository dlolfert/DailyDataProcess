using System;
using System.IO;
using System.Data.SqlClient;

namespace BarChart
{
    class Program
    {
        static private string cs = "Server=Nixon,1466;Database=Barchart;User Id=sa;Password=@a88word";
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            //ImportBarchart();
            
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
                        else if (fields.Length < 3)
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
                            stmt += ("'" + System.DateTime.Now.ToShortDateString() + "')");

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
                        new SqlConnection(cs))
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
                        new SqlConnection(cs))
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