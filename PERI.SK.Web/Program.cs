using HealthChecks.UI.Client;
using Microsoft.AspNetCore.Diagnostics.HealthChecks;
using Microsoft.Extensions.FileProviders;
using PERI.SK.Application.Conversations.Queries;
using PERI.SK.Infrastructure;
using PERI.SK.Infrastructure.AzureOpenAi;
using PERI.SK.Infrastructure.Data;
using PERI.SK.Infrastructure.DeepSeek;
using PERI.SK.Infrastructure.OpenAi;
using PERI.SK.Web.Components;
using PERI.SK.Web;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

builder.Services.ApplyInfrastructureServiceCollection(builder.Configuration);
builder.Services.ApplyDataServiceCollection(builder.Configuration); var aiPlatform = builder.Configuration["AiPlatform:Platform"];

if (aiPlatform == "OpenAI")
{
    builder.Services.ApplyOpenAiServiceCollection(builder.Configuration);
}
else if (aiPlatform == "AzureOpenAI")
{
    builder.Services.ApplyAzureOpenAiServiceCollection(builder.Configuration);
}
else
{
    builder.Services.ApplyDeepSeekServiceCollection(builder.Configuration);
}

builder.Services.AddMediatR(x => x.RegisterServicesFromAssemblyContaining<GetResponseQuery>());

builder.Services.AddMemoryCache();

builder.Services.AddHealthChecks().AddCheck<HealthCheckHandler>(nameof(HealthCheckHandler));

var app = builder.Build();

app.UseStaticFiles();

app.UseStaticFiles(new StaticFileOptions
{
    FileProvider = new PhysicalFileProvider(Path.Combine(Directory.GetCurrentDirectory(), "npm_packages")),
    RequestPath = "/npm_packages"
});

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHealthChecks("/hc", new HealthCheckOptions()
{
    Predicate = _ => true,
    ResponseWriter = UIResponseWriter.WriteHealthCheckUIResponse
});

app.UseHttpsRedirection();

app.UseAntiforgery();

app.MapStaticAssets();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
