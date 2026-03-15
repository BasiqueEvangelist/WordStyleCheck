using Minio;
using WordStyleCheckService.Foreman;
using WordStyleCheckService.Worker;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.Configure<Options>(builder.Configuration);
builder.Services.AddSingleton<Db>();
builder.Services.AddHostedService<CleanupDbService>();

var opts = builder.Configuration.Get<Options>()!;
builder.Services.AddMinio(x => x
    .WithEndpoint(opts.S3EndpointUrl)
    .WithCredentials(opts.S3AccessKeyId, opts.S3SecretAccessKey)
    .WithRegion(opts.S3Region)
    .WithSSL(false));

builder.Services.AddControllers()
    .AddJsonOptions(options => {
        options.JsonSerializerOptions.PropertyNamingPolicy = null;
    });;

var app = builder.Build();

app.UseAuthorization();

app.MapControllers();

app.Run();
