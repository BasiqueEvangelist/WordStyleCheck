using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Minio;
using WordStyleCheckService.Worker;
using Options = WordStyleCheckService.Worker.Options;

var builder = Host.CreateApplicationBuilder(args);

builder.Configuration.AddJsonFile("secrets.json", true);

builder.Services.Configure<Options>(builder.Configuration);
builder.Services.AddSingleton<Db>();

var opts = builder.Configuration.Get<Options>()!;
builder.Services.AddMinio(x => x
        .WithEndpoint(opts.S3EndpointUrl)
        .WithCredentials(opts.S3AccessKeyId, opts.S3SecretAccessKey)
        .WithRegion(opts.S3Region)
        .WithSSL(false));
    
builder.Services.AddHostedService<WorkerService>();

var host = builder.Build();
host.Run();