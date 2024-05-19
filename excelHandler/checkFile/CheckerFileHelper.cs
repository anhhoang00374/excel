using System;
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

                    // Lọc ra các file có đuôi .xlsx, .xlsm, .xltx, .xltm, .xlam, .xlsb (định dạng file Excel hỗ trợ bởi EPPlus)
                    foreach (string file in files)
                    {
                        string extension = Path.GetExtension(file).ToLower();
                        if (extension == ".xlsx" || extension == ".xlsm" || extension == ".xltx" ||
                            extension == ".xltm" || extension == ".xlam" || extension == ".xlsb")
                        {
                            excelFiles.Add(file);
                            Console.WriteLine(file);
                        }
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
    }
}

