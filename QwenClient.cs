using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace QwenSDK
{
    // 数据模型类
    public class QwenMessage
    {
        [JsonProperty("role")]
        public string Role { get; set; }

        [JsonProperty("content")]
        public string Content { get; set; }

        [JsonProperty("tool_calls", NullValueHandling = NullValueHandling.Ignore)]
        public List<ToolCall> ToolCalls { get; set; }

        [JsonProperty("tool_call_id", NullValueHandling = NullValueHandling.Ignore)]
        public string ToolCallId { get; set; }

        [JsonProperty("name", NullValueHandling = NullValueHandling.Ignore)]
        public string Name { get; set; }
    }

    public class ToolCall
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("function")]
        public FunctionCall Function { get; set; }
    }

    public class FunctionCall
    {
        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("arguments")]
        public string Arguments { get; set; }
    }

    public class QwenTool
    {
        [JsonProperty("type")]
        public string Type { get; set; } = "function";

        [JsonProperty("function")]
        public QwenFunction Function { get; set; }
    }

    public class QwenFunction
    {
        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("description")]
        public string Description { get; set; }

        [JsonProperty("parameters")]
        public object Parameters { get; set; }
    }

    public class QwenRequest
    {
        [JsonProperty("model")]
        public string Model { get; set; } = "qwen-plus";

        [JsonProperty("messages")]
        public List<QwenMessage> Messages { get; set; }

        [JsonProperty("tools", NullValueHandling = NullValueHandling.Ignore)]
        public List<QwenTool> Tools { get; set; }

        [JsonProperty("stream", NullValueHandling = NullValueHandling.Ignore)]
        public bool? Stream { get; set; }

        [JsonProperty("temperature", NullValueHandling = NullValueHandling.Ignore)]
        public double? Temperature { get; set; }

        [JsonProperty("max_tokens", NullValueHandling = NullValueHandling.Ignore)]
        public int? MaxTokens { get; set; }
    }

    public class QwenResponse
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("object")]
        public string Object { get; set; }

        [JsonProperty("created")]
        public long Created { get; set; }

        [JsonProperty("model")]
        public string Model { get; set; }

        [JsonProperty("choices")]
        public List<Choice> Choices { get; set; }

        [JsonProperty("usage")]
        public Usage Usage { get; set; }
    }

    public class Choice
    {
        [JsonProperty("index")]
        public int Index { get; set; }

        [JsonProperty("message")]
        public QwenMessage Message { get; set; }

        [JsonProperty("delta")]
        public QwenMessage Delta { get; set; }

        [JsonProperty("finish_reason")]
        public string FinishReason { get; set; }
    }

    public class Usage
    {
        [JsonProperty("prompt_tokens")]
        public int PromptTokens { get; set; }

        [JsonProperty("completion_tokens")]
        public int CompletionTokens { get; set; }

        [JsonProperty("total_tokens")]
        public int TotalTokens { get; set; }
    }

    // 流式响应事件参数
    public class StreamChunkEventArgs : EventArgs
    {
        public QwenResponse Chunk { get; set; }
        public string RawContent { get; set; }
    }

    // 主要的Qwen客户端类
    public class QwenClient : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiKey;
        private const string BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions";

        public event EventHandler<StreamChunkEventArgs> StreamChunkReceived;
        public event EventHandler<Exception> StreamError;
        public event EventHandler StreamCompleted;

        public QwenClient(string apiKey)
        {
            _apiKey = apiKey ?? throw new ArgumentNullException(nameof(apiKey));
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {_apiKey}");
            _httpClient.Timeout = TimeSpan.FromMinutes(5); // 5分钟超时
        }

        /// <summary>
        /// 非流式聊天完成
        /// </summary>
        public async Task<QwenResponse> ChatCompletionAsync(QwenRequest request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request));

            request.Stream = false; // 确保非流式

            try
            {
                var json = JsonConvert.SerializeObject(request, Formatting.None, new JsonSerializerSettings
                {
                    NullValueHandling = NullValueHandling.Ignore
                });

                var content = new StringContent(json, Encoding.UTF8, "application/json");
                var response = await _httpClient.PostAsync(BASE_URL, content);

                var responseContent = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    throw new HttpRequestException($"API请求失败: {response.StatusCode}, {responseContent}");
                }

                return JsonConvert.DeserializeObject<QwenResponse>(responseContent);
            }
            catch (Exception ex)
            {
                throw new Exception($"聊天完成请求失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 流式聊天完成
        /// </summary>
        public async Task ChatCompletionStreamAsync(QwenRequest request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request));

            request.Stream = true; // 确保流式

            try
            {
                var json = JsonConvert.SerializeObject(request, Formatting.None, new JsonSerializerSettings
                {
                    NullValueHandling = NullValueHandling.Ignore
                });

                var content = new StringContent(json, Encoding.UTF8, "application/json");
                var response = await _httpClient.PostAsync(BASE_URL, content);

                if (!response.IsSuccessStatusCode)
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    throw new HttpRequestException($"API请求失败: {response.StatusCode}, {errorContent}");
                }

                using (var stream = await response.Content.ReadAsStreamAsync())
                using (var reader = new StreamReader(stream))
                {
                    string line;
                    while ((line = await reader.ReadLineAsync()) != null)
                    {
                        if (string.IsNullOrWhiteSpace(line))
                            continue;

                        // 处理SSE格式: data: {...}
                        if (line.StartsWith("data: "))
                        {
                            var dataContent = line.Substring(6).Trim();

                            // 检查是否是结束标记
                            if (dataContent == "[DONE]")
                            {
                                StreamCompleted?.Invoke(this, EventArgs.Empty);
                                break;
                            }

                            try
                            {
                                var chunk = JsonConvert.DeserializeObject<QwenResponse>(dataContent);
                                StreamChunkReceived?.Invoke(this, new StreamChunkEventArgs
                                {
                                    Chunk = chunk,
                                    RawContent = dataContent
                                });
                            }
                            catch (JsonException jsonEx)
                            {
                                // 如果JSON解析失败，仍然触发事件但chunk为null
                                StreamChunkReceived?.Invoke(this, new StreamChunkEventArgs
                                {
                                    Chunk = null,
                                    RawContent = dataContent
                                });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                StreamError?.Invoke(this, ex);
                throw new Exception($"流式聊天完成请求失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 带工具调用的聊天完成
        /// </summary>
        public async Task<QwenResponse> ChatCompletionWithToolsAsync(
            List<QwenMessage> messages,
            List<QwenTool> tools,
            string model = "qwen-plus")
        {
            var request = new QwenRequest
            {
                Model = model,
                Messages = messages,
                Tools = tools,
                Stream = false
            };

            return await ChatCompletionAsync(request);
        }

        /// <summary>
        /// 流式工具调用
        /// </summary>
        public async Task ChatCompletionWithToolsStreamAsync(
            List<QwenMessage> messages,
            List<QwenTool> tools,
            string model = "qwen-plus")
        {
            var request = new QwenRequest
            {
                Model = model,
                Messages = messages,
                Tools = tools,
                Stream = true
            };

            await ChatCompletionStreamAsync(request);
        }

        /// <summary>
        /// 便捷方法：简单聊天
        /// </summary>
        public async Task<string> ChatAsync(string userMessage, string systemMessage = null, string model = "qwen-plus")
        {
            var messages = new List<QwenMessage>();

            if (!string.IsNullOrEmpty(systemMessage))
            {
                messages.Add(new QwenMessage { Role = "system", Content = systemMessage });
            }

            messages.Add(new QwenMessage { Role = "user", Content = userMessage });

            var request = new QwenRequest
            {
                Model = model,
                Messages = messages
            };

            var response = await ChatCompletionAsync(request);
            return response.Choices?[0]?.Message?.Content ?? "";
        }

        /// <summary>
        /// 便捷方法：流式聊天，返回完整内容
        /// </summary>
        public async Task<string> ChatStreamAsync(string userMessage, string systemMessage = null,
            Action<string> onChunkReceived = null, string model = "qwen-plus")
        {
            var messages = new List<QwenMessage>();
            var completeContent = new StringBuilder();
            var tcs = new TaskCompletionSource<string>();

            if (!string.IsNullOrEmpty(systemMessage))
            {
                messages.Add(new QwenMessage { Role = "system", Content = systemMessage });
            }

            messages.Add(new QwenMessage { Role = "user", Content = userMessage });

            // 临时订阅事件
            EventHandler<StreamChunkEventArgs> chunkHandler = (sender, e) =>
            {
                if (e.Chunk?.Choices != null && e.Chunk.Choices.Count > 0)
                {
                    var delta = e.Chunk.Choices[0].Delta;
                    if (!string.IsNullOrEmpty(delta?.Content))
                    {
                        completeContent.Append(delta.Content);
                        onChunkReceived?.Invoke(delta.Content);
                    }
                }
            };

            EventHandler completedHandler = (sender, e) =>
            {
                tcs.SetResult(completeContent.ToString());
            };

            EventHandler<Exception> errorHandler = (sender, e) =>
            {
                tcs.SetException(e);
            };

            StreamChunkReceived += chunkHandler;
            StreamCompleted += completedHandler;
            StreamError += errorHandler;

            try
            {
                var request = new QwenRequest
                {
                    Model = model,
                    Messages = messages
                };

                await ChatCompletionStreamAsync(request);
                return await tcs.Task;
            }
            finally
            {
                // 清理事件订阅
                StreamChunkReceived -= chunkHandler;
                StreamCompleted -= completedHandler;
                StreamError -= errorHandler;
            }
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }

    // 工具调用辅助类
    public static class QwenToolBuilder
    {
        public static QwenTool CreateFunction(string name, string description, object parameters = null)
        {
            return new QwenTool
            {
                Type = "function",
                Function = new QwenFunction
                {
                    Name = name,
                    Description = description,
                    Parameters = parameters ?? new { }
                }
            };
        }

        public static QwenTool CreateWeatherTool()
        {
            return CreateFunction(
                "get_current_weather",
                "当你想查询指定城市的天气时非常有用。",
                new
                {
                    type = "object",
                    properties = new
                    {
                        location = new
                        {
                            type = "string",
                            description = "城市或县区，比如北京市、杭州市、余杭区等。"
                        }
                    },
                    required = new[] { "location" }
                }
            );
        }

        public static QwenTool CreateTimeTool()
        {
            return CreateFunction(
                "get_current_time",
                "当你想知道现在的时间时非常有用。",
                new { }
            );
        }
    }
}