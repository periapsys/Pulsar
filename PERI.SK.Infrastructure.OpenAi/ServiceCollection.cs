using PERI.SK.Domain.Interfaces;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.SemanticKernel;

namespace PERI.SK.Infrastructure.OpenAi
{
    public static class ServiceCollection
    {
        public static void ApplyOpenAiServiceCollection(this IServiceCollection services, IConfiguration configuration)
        {
            var model = configuration.GetSection("AiPlatform:Model").Value;
            var apiKey = configuration.GetSection("AiPlatform:ApiKey").Value;

            services.AddSingleton<Kernel>(serviceProvider =>
            {
                var builder = Kernel.CreateBuilder();
                builder.AddOpenAIChatCompletion(model!, apiKey!);
                builder.AddLocalTextEmbeddingGeneration();
                return builder.Build();
            });

            services.AddScoped<IChatService, ChatService>();
        }
    }
}
