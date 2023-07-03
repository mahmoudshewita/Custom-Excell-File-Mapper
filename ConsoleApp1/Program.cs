using ExcellMapper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelMapper<MyData> excelMapper = new ExcelMapper<MyData>();
            string filePath = "department-import-template (6).xlsx";
            int startRow = 1; // Assuming data starts from row 2
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open))
            {
                List<MyData> mappedData = excelMapper.MapExcel(fileStream, startRow);

                // Do something with the mapped data
                foreach (MyData data in mappedData)
                {
                    Console.WriteLine($"Name: {data.Name}, Age: {data.Age}, Count: {data.Count}, Count: {data.Text}");
                }
                //Console.WriteLine($"تجربة");

             }

        }
    }
    public class MyData
    {

        [ExcelColumn(2)]
        public int Count { get; set; }

        [ExcelColumn("A")]
        public string Name { get; set; }

        [ExcelColumn(3)]
        public int Age { get; set; }

        [ExcelColumn("D")]
        public string Text { get; set; }

        [IgnoreExcelColumn]
        public string IgnoreMe { get; set; }
    }
}

