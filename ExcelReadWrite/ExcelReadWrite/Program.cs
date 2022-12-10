using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReadWrite
{
    class Program
    {
        static void Main(string[] args)
        {
            XLS();

            XLSX();

            WriteXLSX();

            // Dừng console lại để xem kq
            Console.ReadKey();

        }

        static void WriteXLSX()
        {
            // Danh sách SV
            List<SinhVien> list = new List<SinhVien>()
            {
                new SinhVien{ MSSV = "15211TT00xx", Name = "Trần Minh Phát", Phone = "090999xxxx" },
                new SinhVien{ MSSV = "15211TT00xx", Name = "Võ Phương Quân", Phone = "090999xxxx" },
                new SinhVien{ MSSV = "15211TT00xx", Name = "Lê Bảo Long", Phone = "090999xxxx" },
                new SinhVien{ MSSV = "15211TT00xx", Name = "Nguyễn Trung Hiếu", Phone = "090999xxxx" },
            };

            // khởi tạo wb rỗng
            XSSFWorkbook wb = new XSSFWorkbook();

            // Tạo ra 1 sheet
            ISheet sheet = wb.CreateSheet();

            // Bắt đầu ghi lên sheet

            // Tạo row
            var row0 = sheet.CreateRow(0);
            // Merge lại row đầu 3 cột
            row0.CreateCell(0); // tạo ra cell trc khi merge
            CellRangeAddress cellMerge = new CellRangeAddress(0, 0, 0, 2);
            sheet.AddMergedRegion(cellMerge);
            row0.GetCell(0).SetCellValue("Thông tin sinh viên");

            // Ghi tên cột ở row 1
            var row1 = sheet.CreateRow(1);
            row1.CreateCell(0).SetCellValue("MSSV");
            row1.CreateCell(1).SetCellValue("Tên");
            row1.CreateCell(2).SetCellValue("Phone");

            // bắt đầu duyệt mảng và ghi tiếp tục
            int rowIndex = 2;
            foreach (var item in list)
            {
                // tao row mới
                var newRow = sheet.CreateRow(rowIndex);

                // set giá trị
                newRow.CreateCell(0).SetCellValue(item.MSSV);
                newRow.CreateCell(1).SetCellValue(item.Name);
                newRow.CreateCell(2).SetCellValue(item.Phone);

                // tăng index
                rowIndex++;
            }

            // xong hết thì save file lại
            FileStream fs = new FileStream(@"C:\Users\ACER\OneDrive\Desktop\Tài Liệu\new1.xlsx", FileMode.CreateNew);
            wb.Write(fs);
        }

        static void XLSX()
        {
            // Lấy stream file
            FileStream fs = new FileStream(@"C: \Users\ACER\Downloads\NPOI_EXCEL\testwb.xlsx", FileMode.Open);

            // Khởi tạo workbook để đọc
            XSSFWorkbook wb = new XSSFWorkbook(fs);

            // Lấy sheet đầu tiên
            ISheet sheet = wb.GetSheetAt(0);

            // đọc sheet này bắt đầu từ row 2 (0, 1 bỏ vì tiêu đề)
            int rowIndex = 2;

            // xuất thong báo
            Console.OutputEncoding = Encoding.UTF8; // để xuất ra console tv có dấu
            Console.WriteLine("Thông tin SV từ file Excel XLSX");

            // nếu vẫn chưa gặp end thì vẫn lấy data
            while (sheet.GetRow(rowIndex).GetCell(0).StringCellValue.ToLower() != "end")
            {
                // lấy row hiện tại
                var nowRow = sheet.GetRow(rowIndex);

                // vì ta dùng 3 cột A, B, C => data của ta sẽ như sau
                var MSSV = nowRow.GetCell(0).StringCellValue;
                var Name = nowRow.GetCell(1).StringCellValue;
                var Phone = nowRow.GetCell(2).StringCellValue;

                // Xuất ra thông tin lên màn hình
                Console.WriteLine("MSSV: {0} | Họ & Tên: {1} | SDT: {2}", MSSV, Name, Phone);

                // tăng index khi lấy xong
                rowIndex++;
            }
        }

        static void XLS()
        {
            // Lấy stream file
            FileStream fs = new FileStream(@"C: \Users\ACER\Downloads\NPOI_EXCEL\testwb.xls", FileMode.Open);

            // Khởi tạo workbook để đọc
            HSSFWorkbook wb = new HSSFWorkbook(fs);

            // Lấy sheet đầu tiên
            ISheet sheet = wb.GetSheetAt(0);

            // đọc sheet này bắt đầu từ row 2 (0, 1 bỏ vì tiêu đề)
            int rowIndex = 2;

            // xuất thong báo
            Console.OutputEncoding = Encoding.UTF8; // để xuất ra console tv có dấu
            Console.WriteLine("Thông tin SV từ file Excel XLS");

            // nếu vẫn chưa gặp end thì vẫn lấy data
            while (sheet.GetRow(rowIndex).GetCell(0).StringCellValue.ToLower() != "end")
            {
                // lấy row hiện tại
                var nowRow = sheet.GetRow(rowIndex);

                // vì ta dùng 3 cột A, B, C => data của ta sẽ như sau
                var MSSV = nowRow.GetCell(0).StringCellValue;
                var Name = nowRow.GetCell(1).StringCellValue;
                var Phone = nowRow.GetCell(2).StringCellValue;

                // Xuất ra thông tin lên màn hình
                Console.WriteLine("MSSV: {0} | Họ & Tên: {1} | SDT: {2}", MSSV, Name, Phone);

                // tăng index khi lấy xong
                rowIndex++;
            }
        }
    }

    public class SinhVien
    {
        public string MSSV;
        public string Name;
        public string Phone;
    }
    
}
