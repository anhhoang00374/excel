using System;
using System.Drawing;
using OfficeOpenXml;

namespace excelHandler.excelCompare
{
	public class CompareHelper
	{
		public CompareHelper()
		{
		}
        

        public List<String> CompareSheet(ExcelWorksheets listSheet, ExcelWorksheets listSheet1)
        {
            List<String> listError = new List<String>();
            for(int index = 0; index < listSheet.Count; index++)
            {
                List<String> errors = compareContentOfSheet(listSheet[index], listSheet1[index]);
                listError.AddRange(errors);
            }
            return listError;
        }

        public List<String> checkListSheet(ExcelWorksheets listSheet, ExcelWorksheets listSheet1)
        {
            List<String> listError = new List<String>();
            List<String> errorName = checkNames(listSheet, listSheet1);
            listError.AddRange(errorName);
            //List<String> errorColor = checkColor(listSheet, listSheet1);
            //listError.AddRange(errorColor);
            return listError;
        }
        

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook1"></param>
        /// <param name="workbook2"></param>
        /// <returns></returns>
        public bool IsEqualNameWorkbook(ExcelWorkbook workbook1, ExcelWorkbook workbook2)
        {
            if (workbook1.Names.Equals(workbook2.Names))
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="errorList"></param>
        public void DisplayError(List<String> errorList)
        {
            Console.WriteLine("count error list: " + errorList.Count);
            foreach(String error in errorList) {
                Console.WriteLine(error);
            }

        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="forder">folder</param>
        /// <param name="fileName"></param>
        /// <param name="isAbsolute"></param>
        /// <returns></returns>
		public FileInfo CreateFileInfo(String forder, String fileName,bool isAbsolute = true)
		{
            var filePath = "";
            if (isAbsolute) {
                filePath = String.Format("{0}/{1}/{2}", GetCurrentDirectory(), forder, fileName);
            }else
            {
                filePath = String.Format("{0}/{1}", forder, fileName);
            }

            return new FileInfo(filePath);
        }

        private List<String> compareContentOfSheet(ExcelWorksheet sheet, ExcelWorksheet sheet1)
        {
            //ExcelCellAddress lastCell = null;
            //lastCell= sheet.Dimension.End;
            List<String> listError = new List<String>();
            int lastRowSheet = sheet.Dimension.End.Row;
            int lastColumnSheet = sheet.Dimension.End.Column;
            int lastRowSheet1 = sheet1.Dimension.End.Row;
            int lastColumnSheet1 = sheet1.Dimension.End.Column;
            if(lastRowSheet != lastRowSheet1)
            {
                listError.Add(String.Format("Sheet({0}): last row is different. value is {1} and {2}", sheet.Name, lastRowSheet, lastRowSheet1));
            }
            if (lastColumnSheet != lastColumnSheet1)
            {
                listError.Add(String.Format("Sheet({0}): last column is different. value is {1} and {2}", sheet.Name, lastColumnSheet, lastColumnSheet1));
            }

            for(int index = 1;index <= lastRowSheet; index++)
            {
                for(int indexColumn = 1; indexColumn <= lastColumnSheet; indexColumn++)
                {
                    List<String> errors = checkCells(sheet, sheet1, index, indexColumn);
                    if(errors.Count != 0)
                    {
                        listError.AddRange(errors);
                    }
                }
            }
            return listError;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns>directory name</returns>
        private String GetCurrentDirectory()
        {
            return Directory.GetCurrentDirectory();
        }

        private List<String> checkNames(ExcelWorksheets listSheet, ExcelWorksheets listSheet1)
        {
            List<String> listError = new List<String>();
            var listNameCheckSheet = listSheet1.Select(sheet => sheet.Name).ToList();

            if (!listSheet.Count.Equals(listSheet1.Count)) {
                listError.Add("number of sheet is not equal");
            } 
            for(int index = 0;index< listSheet.Count;index++)
            {
                if (!listNameCheckSheet.Contains(listSheet[index].Name))
                {
                    listError.Add("sheet name " + listSheet[index].Name + " exits only one file");
                    continue;
                }
                if (!listSheet[index].Name.Equals(listSheet1[index].Name)){
                    listError.Add(String.Format("index of sheet {0} different", listSheet[index].Name));
                }
            }
            return listError;
        }

        private List<String> checkCells(ExcelWorksheet sheet, ExcelWorksheet sheet1,int row,int column)
        {
            List<String> listError = new List<String>();
            var cell = sheet.Cells[row, column];
            var cell1 = sheet1.Cells[row, column];
            //check merge
            
            if((!cell.Merge).Equals(cell1.Merge))
            {
                Console.WriteLine(String.Format("rangre merge: {0},{1}", cell.Address,cell.Offset(0,1).Address));
                listError.Add(String.Format("Cell with row{0} column{1} is different. because one is merged", row, column));
            }

            //checkValue
            
            var isEqualValue = cell.Value?.Equals(cell1.Value);
            var isEqualValue1 = cell1.Value?.Equals(cell.Value);
            bool isNull = (isEqualValue == null && isEqualValue1 == null) ? true: false;
            if (!isNull)
            {
                bool isEqual = Convert.ToBoolean(isEqualValue);
                if (!isEqual)
                {
                    listError.Add(String.Format("Value of cell row {0} column {1} is different", row, column));
                }
            }

            //check border
            bool borderTop = cell.Style.Border.Top.Style == cell1.Style.Border.Top.Style;
            bool borderLeft = cell.Style.Border.Left.Style == cell1.Style.Border.Left.Style;
            bool borderBottom = cell.Style.Border.Bottom.Style == cell1.Style.Border.Bottom.Style;
            bool borderRight = cell.Style.Border.Right.Style == cell1.Style.Border.Right.Style;
            bool checkedBorder = borderTop && borderLeft && borderBottom && borderRight;

            if (!checkedBorder)
            {
                listError.Add(String.Format("Border of cell row {0} column {1} is different", row, column));
            }


            //check align
            

            //{
            //    listError.Add(String.Format("Value of cell with row{0} column{1} is different", row, column));
            //}
            //Console.WriteLine(cell.Merge);
            //Console.WriteLine("sheet 1");
            //Console.WriteLine(cell.Value);
            //Console.WriteLine("sheet 2");
            //Console.WriteLine(cell1.Value);
            //Console.WriteLine(String.Format("result is {0}",cell?.Value?.Equals(cell1.Value)));
            //Console.WriteLine(cell1.Value);
            //if (cell.Value != cell1.Value) {
            //    listError.Add(String.Format("Value at row{0} and column{1} is different", row, column));
            //}
            return listError;
        }

        //private List<String> checkColor(ExcelWorksheets listSheet, ExcelWorksheets listSheet1)
        //{
        //    List<String> listError = new List<String>();
        //    var listNameCheckSheet = listSheet1.Select(sheet => sheet.Name).ToList();

        //    if (!listSheet.Count.Equals(listSheet1.Count))
        //    {
        //        listError.Add("number of sheet is not equal");
        //    }
        //    for (int index = 0; index < listSheet.Count; index++)
        //    {
        //        if (!listNameCheckSheet.Contains(listSheet[index].Name))
        //        {
        //            continue;
        //        }
        //        int sheetColor = listSheet[index].TabColor.ToArgb();
        //        int sheetColor2 = listSheet1[index].TabColor.ToArgb();
        //        Console.WriteLine(sheetColor);
        //        Console.WriteLine(sheetColor2);
        //        if (!sheetColor.Equals(sheetColor2)){
        //            listError.Add(String.Format("color of sheet {0} is different", listSheet[index].Name));
        //            Console.WriteLine(sheetColor);
        //            Console.WriteLine(sheetColor2);
        //        }
        //    }
        //    return listError;
        //}

    }
}

