using WordStyleCheckService.Foreman;
using WordStyleCheckService.Worker;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.Configure<Options>(builder.Configuration);
builder.Services.AddSingleton<Db>();
builder.Services.AddHostedService<CleanupDbService>();

builder.Services.AddControllers();

var app = builder.Build();

app.UseAuthorization();

app.MapControllers();

app.Run();
