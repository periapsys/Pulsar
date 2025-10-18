using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.SemanticKernel.ChatCompletion;
using PERI.SK.Domain.Models;
using PERI.SK.Infrastructure.Data;
using System.Text.RegularExpressions;

namespace PERI.SK.Infrastructure
{
    public partial class ChatService
    {
        /// <summary>
        /// Processes Excel ChatCompletion
        /// </summary>
        /// <param name="refData"></param>
        /// <param name="chatHistory"></param>
        /// <param name="connectionString"></param>
        /// <param name="query"></param>
        /// <returns></returns>
        private async Task<string> ProcessExcel(ReferenceData refData, ChatHistory chatHistory, string connectionString, string? query = null)
        {
            var excelQueries = _serviceProvider.GetRequiredService<ExcelQueries>();

            excelQueries.CreateInMemoryData(connectionString, refData.Subject);

            var subject = refData.Subject;

            var cacheKey = $"{subject}_excel_{nameof(excelQueries.GetFields)}";
            var fields = _cache.Get<string>(cacheKey) ?? await excelQueries.GetFields(connectionString!);
            _cache.Set(cacheKey, fields, new MemoryCacheEntryOptions { AbsoluteExpirationRelativeToNow = TimeSpan.FromHours(24) });
            var prompt = string.Format(await GetPrompt("generate_excel_filter"), fields, refData.Subject, query);

            chatHistory.AddUserMessage(prompt);
            var request = await GetChatResponse(chatHistory);
            chatHistory.AddAssistantMessage(request);

            var pattern = @"```sql(.*?)```";
            var match = Regex.Match(request.ToString(), pattern, RegexOptions.Singleline);
            string data;

            try
            {
                if (match.Success)
                {
                    var sql = match.Groups[1].Value;
                    data = await excelQueries.GetData(connectionString!, sql.Trim().Replace("\n", " "));
                }
                else
                    data = await excelQueries.GetData(connectionString!, request);

                if (string.IsNullOrEmpty(data))
                    return await GetResponse("no_result");

                prompt = string.Format(await GetPrompt("make_data_readable_sql"), data, fields);
                chatHistory.AddUserMessage(prompt);
                request = await GetChatResponse(chatHistory);
                chatHistory.AddAssistantMessage(request);

                return request;
            }
            catch
            {
                return "Unable to process your query.";
            }
        }
    }
}
