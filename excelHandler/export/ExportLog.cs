using System;
namespace export
{
	public class ExportLog
	{
		public ExportLog()
		{
		}
        public void WriteLog(List<String> listLogs)
        {
            String fileName = "logs.txt";
            String currentDir = Directory.GetCurrentDirectory();

            String pathCheck = String.Format("{0}/log/{1}", currentDir, fileName);
            ExportHelper exportHelper = new ExportHelper();
            String path = exportHelper.CreateFile(pathCheck);

            using (var writer = new StreamWriter(path))
            {
                for(var index = 0; index<listLogs.Count; index++)
                {
                    writer.WriteLine(String.Format("error {0}: {1}", index + 1, listLogs[index]));
                }
                
            }
        }
    }
	
}

