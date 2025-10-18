using PERI.SK.Domain.Interfaces;
using PERI.SK.Domain.Models;
using PERI.SK.Infrastructure.Enums;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.ChatCompletion;
using Microsoft.SemanticKernel.Connectors.OpenAI;
using Newtonsoft.Json.Linq;

namespace PERI.SK.Infrastructure
{
    public partial class ChatService : IChatService
    {
        readonly Kernel _kernel;
        readonly Dictionary<string, Func<ReferenceData, ChatHistory, string, string?, Task<string>>> _functionDictionary;
        readonly IReferenceDataService _referenceDataService;
        readonly IMemoryCache _cache;
        readonly IChatCompletionService _chatCompletionService;
        readonly IConfiguration _configuration;
        readonly IServiceProvider _serviceProvider;

        public ChatService(IServiceProvider serviceProvider)
        {
            _kernel = serviceProvider.GetRequiredService<Kernel>();
            _referenceDataService = serviceProvider.GetRequiredService<IReferenceDataService>();
            _chatCompletionService = _kernel.GetRequiredService<IChatCompletionService>();
            _configuration = serviceProvider.GetRequiredService<IConfiguration>();
            _cache = serviceProvider.GetRequiredService<IMemoryCache>();
            _serviceProvider = serviceProvider;

            _functionDictionary = new()
            {
                { nameof(ReferenceDataType.Sql).ToLower(), ProcessSql },
                { nameof(ReferenceDataType.Pdf).ToLower(), ProcessPdf },
                { nameof(ReferenceDataType.Excel).ToLower(), ProcessExcel }
            };
        }

        public async Task<string> Converse(string query, string requestor)
        {
            var subjects = await GetSubjects();

            var chatHistory = GetChatHistory(requestor);

            try
            {
                var prompt = string.Format(await GetPrompt("context"), string.Join(", ", subjects));
                chatHistory.AddSystemMessage(prompt);

                // Construct the prompt asking the assistant to classify the query
                prompt = string.Format(await GetPrompt("is_classified"), string.Join(", ", subjects), query);

                // Add the user message to the chat history
                chatHistory.AddUserMessage(prompt);

                // Get the assistant's response
                var request = await GetChatResponse(chatHistory);

                // Get the result of the assistant's response
                var result = string.Join(" ", request).ToLower();

                // If the assistant's response is "no" (i.e., it's not a sales query), check if it's casual and respond accordingly
                if (result.StartsWith("None", StringComparison.OrdinalIgnoreCase))
                {
                    // We assume if it's not a sales query, it could be a greeting or casual question, 
                    // so we ask the assistant to respond appropriately
                    chatHistory.AddUserMessage(string.Format(await GetPrompt("not_classified"), query));

                    // Get a more appropriate response from the assistant (like a greeting response)
                    var response = await GetChatResponse(chatHistory);
                    chatHistory.AddAssistantMessage(response);
                    return response;
                }

                var refData = await _referenceDataService.GetReferenceData(result);

                result = await _functionDictionary[refData.ReferenceData.Type.ToLower()](refData.ReferenceData, chatHistory, refData.Connectionstring!, query);

                return result;
                
            }
            catch (HttpOperationException ex)
            {
                // TODO: For now, clear history to prevent rate limit
                chatHistory.Clear();
                return $"Please enter your query in 1 min.\n{ex.Message}";
            }
            catch (Exception ex)
            {
                return "Unable to process your query.";
            }
        }

        public async Task<string> GetPrompt(string key)
        {
            return await GetMessage("prompts.json", key);
        }

        public async Task<string> GetResponse(string key)
        {
            return await GetMessage("responses.json", key);
        }

        private async Task<string> GetChatResponse(ChatHistory chatHistory)
        {
            // Now use the limited history to get the response
            var request = await _chatCompletionService.GetChatMessageContentsAsync(
                chatHistory,
                new OpenAIPromptExecutionSettings { MaxTokens = 200 },
                _kernel
            );

            return string.Join(" ", request).Trim();
        }

        private ChatHistory GetChatHistory(string requestor)
        {
            var chatHistory = _cache.Get<ChatHistory?>(requestor) ?? new ChatHistory();
            _cache.Set(requestor, chatHistory, new MemoryCacheEntryOptions { AbsoluteExpirationRelativeToNow = TimeSpan.FromHours(12) });

            return chatHistory;
        }

        private async Task<string> GetMessage(string jsonFile, string key)
        {
            try
            {
                // Parse the JSON file into a JObject
                var config = JObject.Parse(await File.ReadAllTextAsync(jsonFile));

                // Check if the key exists in the JSON and return the associated value
                var value = config[key]?.ToString();

                if (string.IsNullOrEmpty(value))
                {
                    throw new KeyNotFoundException($"Key '{key}' not found in the JSON file.");
                }

                return value;
            }
            catch (FileNotFoundException)
            {
                throw new FileNotFoundException("The prompts.json file was not found.");
            }
            catch (Exception ex)
            {
                throw new Exception($"An error occurred while retrieving the value for key '{key}': {ex.Message}");
            }
        }

        public async Task<List<string>> GetSubjects()
        {
            return _configuration.GetSection("ReferenceData")
                    .AsEnumerable() 
                    .Where(x => x.Key.Contains("Subject"))
                    .Select(x => x.Value)
                    .ToList()!;
        }
    }
}
