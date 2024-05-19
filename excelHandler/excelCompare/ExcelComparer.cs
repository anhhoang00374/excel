using System;
using OfficeOpenXml;
using excelHandler.excelCompare;

namespace excelHandler
{
	public class ExcelComparer
	{
        
		public ExcelComparer()
		{
            
        }

        public List<String> comapreFile()
        {
            CompareHelper helper = new CompareHelper();
            List<String> errorlist = new List<String>();
            
            FileInfo fileInfo = helper.CreateFileInfo("origin", "myworkbook.xlsx");
            FileInfo fileInfo1 = helper.CreateFileInfo("origin", "myworkbook1.xlsx");

            // Kiểm tra xem tệp có tồn tại hay không
            Console.WriteLine(fileInfo.Exists ? "file exits" : "error:  file " + fileInfo.Name + " not fount");
            Console.WriteLine(fileInfo1.Exists ? "file exits": "error:  file " + fileInfo1.Name + " not fount");
            

            if (fileInfo.Exists && fileInfo1.Exists)
            {
                using (var package = new ExcelPackage(fileInfo))
                using (var package1 = new ExcelPackage(fileInfo1))
                {
                    ExcelWorkbook workbook = package.Workbook;
                    ExcelWorkbook workbook1 = package1.Workbook;
                    // check name tow workbooks
                    bool isNameEqual = helper.IsEqualNameWorkbook(workbook, workbook1);
                    if (!isNameEqual)
                    {
                        errorlist.Add("File name is not equal");
                    }

                    ExcelWorksheets sheeetList = workbook.Worksheets;
                    ExcelWorksheets sheeetList1 = workbook1.Worksheets;
                    List<string> errorName = helper.checkListSheet(sheeetList, sheeetList1);
                    errorlist.AddRange(errorName);

                    if(errorlist.Count == 0)
                    {
                        helper.CompareSheet(sheeetList, sheeetList1);
                    }else
                    {
                        List<String> errors = helper.CompareSheet(sheeetList, sheeetList1);
                        errorlist.AddRange(errors);
                        //display list error
                        helper.DisplayError(errorlist);
                    }
                    
                }
            }
            return errorlist;

        }
    }

    
}

