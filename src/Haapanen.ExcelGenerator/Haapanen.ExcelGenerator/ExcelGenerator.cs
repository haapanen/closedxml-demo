using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using ClosedXML.Excel;

namespace Haapanen.ExcelGenerator
{
    public class ExampleClass
    {
        public ExampleClass(string id, string name, string description)
        {
            Id = id;
            Name = name;
            Description = description;
        }

        [Lock]
        public string Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
    }

    public class PropertyMapper<T>
    {

    } 

    public class ExcelGenerator
    {
        public Stream CreateExcel<T>(IEnumerable<T> entities)
            where T : class
        {
            var entityType = typeof(T);

            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.Worksheets.Add(typeof(T).Name);
                sheet.Protect()
                    .SetInsertRows()
                    ;

                var properties = entityType.GetProperties(BindingFlags.Instance | BindingFlags.Public);

                var currentRow = sheet.FirstRow();
                var currentCell = sheet.FirstCell();
                foreach (var property in properties)
                {
                    currentCell.Value = property.Name;
                    currentCell = currentCell.CellRight();
                }

                currentRow = currentRow.RowBelow();
                currentCell = currentRow.FirstCell();
                foreach (var entity in entities)
                {
                    foreach (var property in properties)
                    {
                        if (property.GetCustomAttributes<LockAttribute>().Any())
                        {
                            currentCell.Style.Fill.BackgroundColor = XLColor.Red;
                        }
                        else
                        {
                            currentCell.Style.Protection.SetLocked(false);
                        }

                        currentCell.Value = property.GetValue(entity);
                        currentCell = currentCell.CellRight();
                    }

                    currentRow = currentRow.RowBelow();
                    currentCell = currentRow.FirstCell();
                }

                foreach (var value in Enumerable.Range(0, 500))
                {
                    foreach (var property in properties)
                    {
                        if (property.GetCustomAttributes<LockAttribute>().Any())
                        {
                            currentCell.Style.Fill.BackgroundColor = XLColor.Red;
                        }
                        else
                        {
                            currentCell.Style.Protection.SetLocked(false);
                        }

                        currentCell = currentCell.CellRight();
                    }

                    currentRow = currentRow.RowBelow();
                    currentCell = currentRow.FirstCell();
                }

                sheet.Columns().AdjustToContents();
                var stream = new MemoryStream();
                workbook.SaveAs(stream);
                stream.Position = 0;
                return stream;
            }
        }
    }
}
