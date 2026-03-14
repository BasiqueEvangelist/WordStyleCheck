using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using WordStyleCheckService.Worker;

var builder = Host.CreateApplicationBuilder(args);

builder.Services.Configure<Options>(builder.Configuration);
builder.Services.AddSingleton<Db>();
builder.Services.AddHostedService<WorkerService>();

var host = builder.Build();
host.Run();