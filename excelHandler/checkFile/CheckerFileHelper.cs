using System;
using OfficeOpenXml;

namespace checkFileHelper
{
	public class CheckerFileHelper
	{
		public CheckerFileHelper()
		{
		}
        public List<String> getAllFile(String path)
        {
            List<string> excelFiles = new List<string>();

            try
            {
                // Kiểm tra xem đường dẫn thư mục có tồn tại hay không
                if (Directory.Exists(path))
                {
                    // Lấy danh sách tất cả các file trong thư mục
                    string[] files = Directory.GetFiles(path);
                    String[] d = Directory.GetDirectories(path);
                    // Lọc ra các file có đuôi .xlsx, .xlsm, .xltx, .xltm, .xlam, .xlsb (định dạng file Excel hỗ trợ bởi EPPlus)
                    foreach (string file in files)
                    {
                        string extension = Path.GetExtension(file).ToLower();
                        if (extension == ".xlsx" || extension == ".xlsm" || extension == ".xltx" ||
                            extension == ".xltm" || extension == ".xlam" || extension == ".xlsb")
                        {
                            excelFiles.Add(file);
                            
                        }
                        Console.WriteLine(Path.GetExtension(file).ToLower());
                        Console.WriteLine(Path.GetFileName(file));
                    }
                    foreach (string file in d)
                    {
                        
                        Console.WriteLine(file);
                    }
                }
                else
                {
                    // Nếu đường dẫn thư mục không tồn tại, trả về danh sách rỗng
                    Console.WriteLine($"Đường dẫn thư mục '{path}' không tồn tại.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi: {ex.Message}");
            }

            return excelFiles;
        }

        internal void checkFile(List<string> listFile)
        {
            foreach(var file in listFile)
            {
                using (var package = new ExcelPackage(file))
                {
                    ExcelWorkbook workbook = package.Workbook;
                    ExcelWorksheets worksheets = workbook.Worksheets;
                    foreach(var sheet in worksheets)
                    {
                        if(sheet.Name != null)
                        {
                            checkSheet(sheet);
                        }
                    }
                    package.Save();
                }
            }
            
        }
        private List<String> checkSheet(ExcelWorksheet excelWorksheet)
        {
            var row = excelWorksheet.Column(1);
            row.Width = 100;
            return null;
        }
    }
}

