using System;
using System.Collections.Generic;
using System.IO;

namespace Haapanen.ExcelGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var s = new ExcelGenerator().CreateExcel(new List<ExampleClass>
            {
                new ExampleClass("id1", "name1", "desc1"),
                new ExampleClass("id2", "name2", "desc2"),
                new ExampleClass("id3", "name3", "desc3")
            }))
            {
                using (var fs = File.OpenWrite("output.xlsx"))
                {
                    s.CopyTo(fs);
                    fs.Flush();
                }
            }
        }
    }
}
