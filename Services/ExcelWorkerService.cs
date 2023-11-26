using WorkerServiceExcel.Utils;

namespace WorkerServiceExcel.Services
{
    internal class ExcelWorkerService : BackgroundService
    {
        protected override Task ExecuteAsync(CancellationToken stoppingToken)
        {
            string folderPath = @"C:\Users\Hector\Downloads";
            ExcelFile excelFile = new ExcelFile(folderPath);
            excelFile.GenerateFile();
            return Task.CompletedTask;
        }
    }
}
