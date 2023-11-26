using WorkerServiceExcel;
using WorkerServiceExcel.Services;

IHost host = Host.CreateDefaultBuilder(args)
    .ConfigureServices(services =>
    {
        //services.AddHostedService<Worker>();
        services.AddHostedService<ExcelWorkerService>();
    })
    .Build();

await host.RunAsync();
