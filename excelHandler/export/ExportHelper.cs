using System;
namespace export
{
	public class ExportHelper
	{
		public ExportHelper()
		{
		}
        public String CreateFile(String path)
        {
            FileInfo fileInfo = new FileInfo(path);
            if (fileInfo.Exists)
            {
                Console.WriteLine("File is exits.");
            }else
            {
                FileStream fs = File.Create(path);
                Console.WriteLine("File is not found.");
                Console.WriteLine("File has created.");
            }
            return path;
        }
    }
}

