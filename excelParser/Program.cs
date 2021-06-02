
using System;
using System.Data;
using System.Data.OleDb;
using System.Windows;

namespace excelParser
{
    class Program
    {
        static void Main(string[] args)
        {
            //Console.WriteLine("Hello World!");


            string filePath = @"C:\Users\Serik_Seidigalimov\Desktop\Nike\Coverage Report_SU21_FA21_HO21_ 17.05.2021 Inline Partner Digital Global.xlsx";
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source='" + filePath + "';Extended Properties=\"Excel 12.0;HDR=YES;\"";

            string sheetName = "Detail";
            string sqlScript = "select [Prod cd], [Styl Nm], [Cust Sold To Cd], [Revised task date], sum([Opn To Ref Qty]) from [" + sheetName + "$] group by [Prod cd], [Styl Nm], [Cust Sold To Cd], [Revised task date] having [Revised task date] > 2000 and ([Cust Sold To Cd] = '0008000028' or [Cust Sold To Cd] = '0000406920')";           
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {

                try
                {
                    conn.Open();
                    // Connect EXCEL sheet with OLEDB using connection string
                    OleDbDataAdapter objDA = new System.Data.OleDb.OleDbDataAdapter(sqlScript, conn);
                    DataSet excelDataSet = new DataSet();
                    objDA.Fill(excelDataSet);
                    Console.WriteLine("worked");
                    DataTable excelDataTable = excelDataSet.Tables[0];
                    conn.Close();

                    Console.WriteLine("Row count: " + excelDataTable.Rows.Count);
                    foreach (DataRow row in excelDataTable.Rows) 
                    {
                        foreach (DataColumn column in excelDataTable.Columns)
                        {
                            Console.Write(row[column.ColumnName].ToString() + " ");
                        }
                        Console.WriteLine();
                        break;
                    }

                }
                catch (Exception ex) {
                    Console.WriteLine(ex.Message);                    
                }

                Console.ReadKey();
            }


        }
    }
}
