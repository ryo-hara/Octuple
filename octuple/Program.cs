using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace octuple
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = new XLWorkbook("test.xlsx");
            var worksheet = workbook.Worksheet(1);

            var cell = worksheet.Cell(1, 2);
            cell.Value = "Test";

            workbook.Save();

            return;
        }
    }
}
