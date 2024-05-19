﻿using System;
using checkFileHelper;
namespace checkFile
{
	public class ExcelFileCheck
	{
		public ExcelFileCheck()
		{
		}
        public List<String> Check()
        {
			CheckerFileHelper checkerFileHelper = new CheckerFileHelper();
			String directory = String.Format("{0}/{1}", Directory.GetCurrentDirectory(), "origin");
			//get list file
			List<String> listFile = checkerFileHelper.getAllFile(directory);
			Console.WriteLine(String.Format("number of files: {0}", listFile.Count));
			//check  file
			checkerFileHelper.checkFile(listFile);
            return null;
        }
    }
	 
}

