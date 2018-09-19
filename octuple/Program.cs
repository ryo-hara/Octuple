using System;
using System.IO;
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

        public const string DEFAULT_CONVERT_FILE = "test.xlsx";
        public const string DEFAULT_LIST_FOLDER = "DataFolder\\";

        public const string KEY_VALUE_HIGH_SCHOOL = "99";
        public const string KEY_VALUE_JUNIOR_HIGH_SCHOOL = "88";

        public const int REFERENCE_NUMBER_DIGIT = 6;//整理番号の桁数

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
            XLWorkbook workbook = new XLWorkbook(Constants.DEFAULT_CONVERT_FILE);
            IXLWorksheet worksheet = workbook.Worksheet(1);

            for (int i = Constants.MIN_ROW_NUM; i < Constants.MAX_ROW_NUM; i++)
            {
                for (int n = Constants.MIN_COLUMN_NUM; n < Constants.MAX_COLUMN_NUM; n++)
                {
                    IXLCell _cell = worksheet.Cell(i, n);
                    ConvertCell(ref _cell);
                    worksheet.Cell(i, n).Value = _cell.Value;
                }
                Console.Write("\n");
            }

            workbook.Save();

            DisplayCell(worksheet, Constants.MAX_ROW_NUM, Constants.MAX_COLUMN_NUM);

            return;
        }

        private void ConvertCell(ref IXLCell _cell)
        {
            string _cell_value = _cell.Value.ToString();
            string _parse_cell_value;
            int number_for_lead;

            number_for_lead = _cell_value.IndexOf("99");

            if (number_for_lead >= 0)
            {
                _parse_cell_value = _cell_value.Substring(number_for_lead, Constants.REFERENCE_NUMBER_DIGIT);
            }

            number_for_lead = _cell_value.IndexOf("88");

            if (number_for_lead >= 0)
            {
                _parse_cell_value = _cell_value.Substring(number_for_lead, Constants.REFERENCE_NUMBER_DIGIT);
            }

            /*ファイル読み込み系処理*/

            string _file_name = _parse_cell_value.Substring(0,4)+".txt";//881122
            string _data_file_directory = Constants.DEFAULT_LIST_FOLDER;
            int _attendanc_number = _parse_cell_value.Substring(4,2);//出席番号

            StreamReader _reading_data_file = new StreamReader(_data_file_directory + _file_name, System.Text.Encoding.GetEncoding(932)/*文字コードを指定*/);//テキストファイルのオープン

            string _read_name = "";
            int _read_line_num = 0;

            while (_read_name != null)
            {
                _read_name = _reading_data_file.ReadLine();//テキストファイルから読みだした一行を変数に保存
                if (_read_line_num == _attendanc_number)
                {
                    break;
                }
                _read_line_num++;
            }
            _reading_data_file.Close();

            _cell.Value = _read_name;

            return;
        }

        private void DisplayCell(IXLWorksheet _worksheet ,int _row_range, int _column_range)
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
