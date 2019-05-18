using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LargeDataInsertFromExcel
{
    class Program
    {
        //ok
        static void Main(string[] args)
        {
            var dt = ReaderSheet.Run(@"C:\Users\tiago\OneDrive\problema.xlsx");
            var Inumerable = dt.First().AsEnumerable().ToChunks(100000)
                 .Select(rows => rows.CopyToDataTable());

            foreach (var item in Inumerable)
            {
                Bulk(item);
            }
        }

        public static void Bulk(DataTable dataTable) {

            string connectionString = "Server=mydb.coyr4gzqqdng.us-east-1.rds.amazonaws.com;Database=LongShortDB; User ID=;Password=";
            using (SqlConnection connection =
                new SqlConnection(connectionString))
            {
                // make sure to enable triggers
                // more on triggers in next post
                SqlBulkCopy bulkCopy =
                    new SqlBulkCopy
                    (
                        connection,
                        SqlBulkCopyOptions.TableLock |
                        SqlBulkCopyOptions.FireTriggers |
                        SqlBulkCopyOptions.UseInternalTransaction,
                    null
                    );

                // set the destination table name
                bulkCopy.DestinationTableName = "dbo.TESTE";
                connection.Open();

                // write the data in the "dataTable"
                bulkCopy.WriteToServer(dataTable);
                connection.Close();
            }
            dataTable.Clear();
        }


    }

    public static class ReaderSheet
    {
        public static List<DataTable> Run(string path)
        {
            using (var excel = new ExcelPackage())
            {
                using (var stream = new StreamReader(path))
                {
                    excel.Load(stream.BaseStream);
                    var worksheets = excel.Workbook.Worksheets;
                    var listDataTable = new List<DataTable>();
                    DataTable dataTable = null;

                    foreach (var worksheet in worksheets)
                    {
                        var dimensions = worksheet.Dimension;
                        var properties = new List<string>();
                        dataTable = new DataTable();

                        for (int row = 1; row <= dimensions.Rows; row++)
                        {
                            var itensRow = new List<Object>();
                            for (int col = 1; col <= dimensions.Columns; col++)
                            {
                                var cell = worksheet.Cells[row, col].Value.ToString();
                                if (row == 1)
                                {//creating colluns data table, just for first row
                                    dataTable.Columns.Add(cell, typeof(String));
                                }
                                else
                                {
                                    itensRow.Add(cell);
                                }
                                // Console.WriteLine(cell);
                            }
                            dataTable.Rows.Add(itensRow.ToArray());
                        }
                        listDataTable.Add(dataTable);
                    }
                    return listDataTable;
                }
            }
        }
    }

    public static class help {

        public static IEnumerable<IEnumerable<T>> ToChunks<T>(this IEnumerable<T> enumerable,
                                                     int chunkSize)
        {
            int itemsReturned = 0;
            var list = enumerable.ToList(); // Prevent multiple execution of IEnumerable.
            int count = list.Count;
            while (itemsReturned < count)
            {
                int currentChunkSize = Math.Min(chunkSize, count - itemsReturned);
                yield return list.GetRange(itemsReturned, currentChunkSize);
                itemsReturned += currentChunkSize;
            }
        }

    }
}
