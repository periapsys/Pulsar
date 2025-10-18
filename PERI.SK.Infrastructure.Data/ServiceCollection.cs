using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace PERI.SK.Infrastructure.Data
{
    public static class ServiceCollection
    {
        public static void ApplyDataServiceCollection(this IServiceCollection services, IConfiguration configuration)
        {
            services.AddScoped<SqlQueries>();
            services.AddScoped<PdfQueries>();
            services.AddScoped<ExcelQueries>();
        }
    }
}
