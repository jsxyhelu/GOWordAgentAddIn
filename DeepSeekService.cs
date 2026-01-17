using System;
using System.Net.Http;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace GOWordAgentAddIn
{
    /// <summary>
    /// DeepSeek API 调用服务
    /// </summary>
    public class DeepSeekService
    {
        private readonly string _apiKey;
        private readonly string _apiUrl = "https://api.deepseek.com/v1/chat/completions";
        private readonly HttpClient _httpClient;

        public DeepSeekService(string apiKey)
        {
            _apiKey = apiKey;

            var handler = new HttpClientHandler
            {
                SslProtocols = System.Security.Authentication.SslProtocols.Tls12
            };
            _httpClient = new HttpClient(handler);
            _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {_apiKey}");
        }

        /// <summary>
        /// 发送消息到 DeepSeek 并获取响应
        /// </summary>
        /// <param name="userMessage">用户消息</param>
        /// <param name="model">模型名称，默认为 deepseek-chat</param>
        /// <returns>AI 回复内容</returns>
        public async Task<string> SendMessageAsync(string userMessage, string model = "deepseek-chat")
        {
            try
            {
                var requestBody = new
                {
                    model = model,
                    messages = new[]
                    {
                        new
                        {
                            role = "user",
                            content = userMessage
                        }
                    },
                    temperature = 0.7,
                    max_tokens = 2000
                };

                string jsonContent = JsonConvert.SerializeObject(requestBody);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                HttpResponseMessage response = await _httpClient.PostAsync(_apiUrl, content);
                string responseBody = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    JObject jsonResponse = JObject.Parse(responseBody);
                    var reply = jsonResponse["choices"]?[0]?["message"]?["content"]?.ToString();
                    return reply ?? "未获取到回复内容";
                }
                else
                {
                    return $"API 调用失败: {response.StatusCode}\n{responseBody}";
                }
            }
            catch (Exception ex)
            {
                return $"发生错误: {ex.Message}";
            }
        }

        /// <summary>
        /// 发送带历史记录的消息
        /// </summary>
        /// <param name="messages">消息历史记录</param>
        /// <param name="model">模型名称</param>
        /// <returns>AI 回复内容</returns>
        public async Task<string> SendMessagesWithHistoryAsync(object[] messages, string model = "deepseek-chat")
        {
            try
            {
                var requestBody = new
                {
                    model = model,
                    messages = messages,
                    temperature = 0.7,
                    max_tokens = 2000
                };

                string jsonContent = JsonConvert.SerializeObject(requestBody);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                HttpResponseMessage response = await _httpClient.PostAsync(_apiUrl, content);
                string responseBody = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    JObject jsonResponse = JObject.Parse(responseBody);
                    var reply = jsonResponse["choices"]?[0]?["message"]?["content"]?.ToString();
                    return reply ?? "未获取到回复内容";
                }
                else
                {
                    return $"API 调用失败: {response.StatusCode}\n{responseBody}";
                }
            }
            catch (Exception ex)
            {
                return $"发生错误: {ex.Message}";
            }
        }
    }
}
