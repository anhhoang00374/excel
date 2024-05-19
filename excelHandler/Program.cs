using excelHandler;
using OfficeOpenXml;
using export;
class Program
{
    static void Main(string[] args)
    {



        ExcelPackage.LicenseContext = LicenseContext.Commercial;
        //program execution starts from here
        Console.WriteLine("Command line Arguments: {0}");
        ExcelComparer comparer = new ExcelComparer();

        List<String> errorList = comparer.comapreFile();
        ExportLog exportLog = new ExportLog();
        exportLog.WriteLog(errorList);
    }
}