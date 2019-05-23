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
            var cabecalho = new Dictionary<string, object>()
            {
                {"NOME", typeof(String) },
                {"VALOR", typeof(decimal) },
                {"QUANTIDADE", typeof(Int32) },
                {"DATA", typeof(DateTime) },
                {"OBS", typeof(string) },
            };

            var dt = ReaderSheet.Run(@"C:\Projetos\Comparacao de arquivos exemplos\SAS.xlsx", cabecalho);

            //foreach (DataRow row in dt.First().Rows)
            //{
            //    // ... Write value of first field as integer.
            //    Console.WriteLine(row.Field<DateTime?>(1));
            //}

            var Inumerable = dt.First().AsEnumerable().ToChunks(100000)
                 .Select(rows => rows.CopyToDataTable());

            foreach (var item in Inumerable)
            {
               // var data = item.Rows[1]["DATA"].ToString();

                Bulk(item);
            }

            Console.WriteLine("Done!");
        }

        public static void Bulk(DataTable dataTable) {

            string connectionString = "Server=;Database=TESTE; User ID=;Password=";
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
                bulkCopy.DestinationTableName = "dbo.SAS";
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
        public static List<DataTable> Run(string path, Dictionary<string, dynamic> cabecalho)
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
                                var cell = worksheet.Cells[row, col].Value?.ToString();
                                if (row == 1)
                                {//creating colluns data table, just for first row

                                    var tipo = cabecalho.Where( x=> x.Key == cell ).First();

                                    //   itensRow.Add(new );

                                      dataTable.Columns.Add(cell, tipo.Value);
                                }
                                else
                                {
                                    itensRow.Add(cell);
                                }
                                // Console.WriteLine(cell);
                            }

                            if (itensRow.Count>0)
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
