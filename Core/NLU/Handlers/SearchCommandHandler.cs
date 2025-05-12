using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using NanoAI.Gemini;

namespace NanoAI.Core.NLU.Handlers
{
    /// <summary>
    /// Handles search commands for retrieving information from the web
    /// </summary>
    public class SearchCommandHandler : ICommandHandler
    {
        private readonly GeminiService _geminiService;
        private readonly HttpClient _httpClient;
        
        // Google Search API endpoint for free custom search
        private const string SEARCH_API_URL = "https://www.googleapis.com/customsearch/v1";
        private const string SEARCH_ENGINE_ID = "YOUR_SEARCH_ENGINE_ID"; // Replace with your own
        
        // Prompt for generating answers from search results
        private const string ANSWER_PROMPT = @"
The user asked: ""{0}""

I found these search results:
{1}

Based solely on the information in these search results, please provide a concise answer.
If the search results don't contain the answer, say that you can't answer based on the available information.
DO NOT use your own knowledge outside of what's in the search results.
Answer in simple, clear language. Start your response directly with the answer, no introduction needed.
";

        public string CommandType => "search";

        public SearchCommandHandler(GeminiService geminiService)
        {
            _geminiService = geminiService ?? throw new ArgumentNullException(nameof(geminiService));
            _httpClient = new HttpClient();
        }

        public bool CanHandle(GeminiCommand command)
        {
            return command.CommandType.Equals("search", StringComparison.OrdinalIgnoreCase);
        }

        public async Task<CommandResult> ExecuteAsync(GeminiCommand command)
        {
            string query = command.Target;
            
            if (string.IsNullOrEmpty(query))
            {
                return new CommandResult
                {
                    Success = false,
                    Message = "No search query provided",
                    Suggestions = new List<string> { "Please provide a search query" }
                };
            }
            
            try
            {
                // First try to get an answer directly from Gemini
                var directAnswer = await GetAnswerFromGemini(query);
                
                if (!string.IsNullOrEmpty(directAnswer))
                {
                    return new CommandResult
                    {
                        Success = true,
                        Message = directAnswer
                    };
                }
                
                // If Gemini doesn't have a direct answer, perform a web search
                var searchResults = await PerformWebSearch(query);
                
                if (searchResults.Count == 0)
                {
                    return new CommandResult
                    {
                        Success = false,
                        Message = "No search results found",
                        Suggestions = new List<string> { "Try a different search query" }
                    };
                }
                
                // Format search results for display
                StringBuilder resultBuilder = new StringBuilder();
                resultBuilder.AppendLine($"Search results for: {query}");
                resultBuilder.AppendLine();
                
                foreach (var result in searchResults)
                {
                    resultBuilder.AppendLine($"- {result.Title}");
                    resultBuilder.AppendLine($"  {result.Snippet}");
                    resultBuilder.AppendLine($"  {result.Link}");
                    resultBuilder.AppendLine();
                }
                
                // Try to generate an answer from the search results
                string answer = await GenerateAnswerFromSearchResults(query, searchResults);
                
                if (!string.IsNullOrEmpty(answer))
                {
                    resultBuilder.AppendLine("Answer:");
                    resultBuilder.AppendLine(answer);
                }
                
                return new CommandResult
                {
                    Success = true,
                    Message = resultBuilder.ToString()
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Search error: {ex.Message}");
                return new CommandResult
                {
                    Success = false,
                    Message = $"Search error: {ex.Message}",
                    Suggestions = new List<string> { "Please try a different search query" }
                };
            }
        }
        
        /// <summary>
        /// Gets a direct answer from Gemini AI
        /// </summary>
        private async Task<string> GetAnswerFromGemini(string query)
        {
            try
            {
                // Send a simple prompt to Gemini to see if it can answer directly
                string prompt = $"The user asked: \"{query}\". If you know the factual answer, provide it concisely. If it requires real-time information you don't have, respond with NEED_SEARCH.";
                
                string response = await _geminiService.SendRawPromptAsync(prompt);
                
                // If the answer indicates we need to search, return empty
                if (response.Contains("NEED_SEARCH") || 
                    response.Contains("I don't have access to real-time information") ||
                    response.Contains("I can't provide real-time information") ||
                    response.Contains("my training data only goes up to"))
                {
                    return string.Empty;
                }
                
                return response;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Gemini answer error: {ex.Message}");
                return string.Empty;
            }
        }
        
        /// <summary>
        /// Performs a web search using Google Search API
        /// </summary>
        private async Task<List<SearchResult>> PerformWebSearch(string query)
        {
            var results = new List<SearchResult>();
            
            try
            {
                // Use browser fallback for now, as we don't have a real API key
                Console.WriteLine($"Performing web search for: {query}");
                
                // In a real implementation, you would use:
                // string url = $"{SEARCH_API_URL}?key={_geminiService._apiKey}&cx={SEARCH_ENGINE_ID}&q={Uri.EscapeDataString(query)}";
                // var response = await _httpClient.GetStringAsync(url);
                // var searchResponse = JsonSerializer.Deserialize<GoogleSearchResponse>(response);
                
                // For now, we'll just create some dummy results
                results.Add(new SearchResult
                {
                    Title = "Search result 1 for " + query,
                    Snippet = "This is a snippet of the search result containing information about " + query,
                    Link = "https://example.com/result1"
                });
                
                results.Add(new SearchResult
                {
                    Title = "Search result 2 for " + query,
                    Snippet = "Another snippet with different information about " + query,
                    Link = "https://example.com/result2"
                });
                
                results.Add(new SearchResult
                {
                    Title = "Search result 3 for " + query,
                    Snippet = "A third snippet with more details about " + query,
                    Link = "https://example.com/result3"
                });
                
                return results;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Web search API error: {ex.Message}");
                return results;
            }
        }
        
        /// <summary>
        /// Generates an answer from search results using Gemini
        /// </summary>
        private async Task<string> GenerateAnswerFromSearchResults(string query, List<SearchResult> searchResults)
        {
            try
            {
                // Format search results as text
                var resultsText = new StringBuilder();
                
                foreach (var result in searchResults)
                {
                    resultsText.AppendLine($"Title: {result.Title}");
                    resultsText.AppendLine($"Snippet: {result.Snippet}");
                    resultsText.AppendLine($"URL: {result.Link}");
                    resultsText.AppendLine();
                }
                
                // Create prompt with query and search results
                string prompt = string.Format(ANSWER_PROMPT, query, resultsText.ToString());
                
                // Get answer from Gemini
                string answer = await _geminiService.SendRawPromptAsync(prompt);
                
                return answer;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Answer generation error: {ex.Message}");
                return string.Empty;
            }
        }
    }

    /// <summary>
    /// Represents a single search result
    /// </summary>
    public class SearchResult
    {
        public string Title { get; set; }
        public string Snippet { get; set; }
        public string Link { get; set; }
    }
} 