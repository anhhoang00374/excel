using excelHandler;
using OfficeOpenXml;
using export;
using checkFile;
class Program
{
    static void Main(string[] args)
    {
        //add license
        ExcelPackage.LicenseContext = LicenseContext.Commercial;

        //check file
        ExcelFileCheck excelFileCheck = new ExcelFileCheck();
        excelFileCheck.Check();

        //
        
        //program execution starts from here
        //Console.WriteLine("Command line Arguments: {0}");
        //ExcelComparer comparer = new ExcelComparer();

        //List<String> errorList = comparer.comapreFile();
        //ExportLog exportLog = new ExportLog();
        //exportLog.WriteLog(errorList);
    }
}