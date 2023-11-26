using OfficeOpenXml;
using WorkerServiceExcel.Entities;

namespace WorkerServiceExcel.Utils
{
    internal class ExcelFile
    {
        private FileSystemWatcher? watcher;
        private string FolterPath;

        public ExcelFile(string FolterPath)
        {
            this.FolterPath = FolterPath;
        }

        
        internal Task GenerateFile()
        {
            watcher = new FileSystemWatcher(FolterPath);
            watcher.NotifyFilter = NotifyFilters.FileName;
            watcher.Created += OnCreated;
            watcher.EnableRaisingEvents = true;
            return Task.CompletedTask;
        }

        private void OnCreated(object sender, FileSystemEventArgs e)
        {
            string[] excelFiles = Directory.GetFiles(FolterPath, "*.xlsx");

            foreach (string filePath in excelFiles)
            {
                // Abrir libro de Excel
                using (ExcelPackage package = new(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                        {
                            if (worksheet.Cells[row, col].Value != null)
                            {
                                string? value = worksheet.Cells[row, col].Value.ToString();
                                string fileName = filePath.Replace(FolterPath, "");
                                var (isNumber, number) = IsNumber(value);

                                if (isNumber)
                                {
                                    var parameterWrite = new ParameterWrite<int>()
                                    {
                                        Col = col,
                                        Row = row,
                                        Path = FolterPath + @"\templates" + fileName,
                                        Value = number
                                    };

                                    Write(parameterWrite);
                                }

                                else
                                {
                                    var parameterWrite = new ParameterWrite<string>()
                                    {
                                        Col = col,
                                        Row = row,
                                        Path = FolterPath + @"\templates" + fileName,
                                        Value = value
                                    };

                                    Write(parameterWrite);
                                }
                            }
                        }
                    }
                }
            }
        }

        private (bool, int) IsNumber(string value)
        {
            int number;
            return (int.TryParse(value, out number), number);
        }

        private void Write(ParameterWrite<string> parameter)
        {
            using (ExcelPackage package = new(new FileInfo(parameter.Path)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                worksheet.Cells[parameter.Row, parameter.Col].Value = parameter.Value;
                package.Save();
            }
        }

        private void Write(ParameterWrite<int> parameter)
        {
            using (ExcelPackage package = new(new FileInfo(parameter.Path)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                worksheet.Cells[parameter.Row, parameter.Col].Value = parameter.Value;
                package.Save();
            }
        }
    }
}
