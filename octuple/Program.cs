using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace octuple
{

    class Constants
    {
        public const int MIN_ROW_NUM = 1;      //シート内で扱う最小の行
        public const int MIN_COLUMN_NUM = 1;   //シート内で扱う最小の行

        public const int MAX_ROW_NUM = 100;    //シート内で扱う最大の行
        public const int MAX_COLUMN_NUM = 100; //シート内で扱う最大の列
    }


    class Program
    {
        static void Main(string[] args)
        {
            FileController file_controller = new FileController();
            file_controller.ConversionWorkbook();

            return;
        }

    }


    class FileController
    {

        public void ConversionWorkbook()
        {
            XLWorkbook workbook = new XLWorkbook("test.xlsx");
            IXLWorksheet worksheet = workbook.Worksheet(1);

            var cell = worksheet.Cell(1, 2);
            cell.Value = "Test";

            workbook.Save();

            Console.WriteLine(worksheet.Cell(1, 2).Value.ToString());

            DisplayCell(worksheet, Constants.MAX_ROW_NUM, Constants.MAX_COLUMN_NUM);

            return;
        }

        public void DisplayCell(IXLWorksheet _worksheet ,int _row_range, int _column_range)
        {
            for (int i = Constants.MIN_ROW_NUM; i < _row_range; i++)
            {
                for (int n = Constants.MIN_COLUMN_NUM; n < _column_range; n++)
                {
                    Console.Write(_worksheet.Cell(i, n).Value.ToString() + ",");
                }
                Console.Write("\n");
            }
            return;
        }

    }
}
