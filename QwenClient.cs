using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net.Http;
using System.Text;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using System;
using System.Linq;

namespace FunctionCallingAI
{
    // 枚举定义
    public enum ChatFinishReason
    {
        Stop,
        ToolCalls,
        Length,
        ContentFilter
    }

    public enum MessageRole
    {
        System,
        User,
        Assistant,
        Tool
    }

    // 消息内容类
    public class MessageContent
    {
        public string Text { get; set; }

        public MessageContent(string text)
        {
            Text = text;
        }
    }

    // 基础消息类
    public abstract class ChatMessage
    {
        public abstract MessageRole Role { get; }
        public abstract List<MessageContent> Content { get; }

        public virtual string ToJsonRole()
        {
            return Role.ToString().ToLower();
        }
    }

    // 用户消息
    public class UserChatMessage : ChatMessage
    {
        public override MessageRole Role { get { return MessageRole.User; } }
        public override List<MessageContent> Content { get; }

        public UserChatMessage(string content)
        {
            Content = new List<MessageContent> { new MessageContent(content) };
        }
    }

    // 系统消息
    public class SystemChatMessage : ChatMessage
    {
        public override MessageRole Role { get { return MessageRole.System; } }
        public override List<MessageContent> Content { get; }

        public SystemChatMessage(string content)
        {
            Content = new List<MessageContent> { new MessageContent(content) };
        }
    }

    // 工具调用信息
    public class ChatToolCall
    {
        public string Id { get; set; }
        public string Type { get; set; } = "function";
        public string FunctionName { get; set; }
        public string FunctionArguments { get; set; }
    }

    // 助手消息
    public class AssistantChatMessage : ChatMessage
    {
        public override MessageRole Role { get { return MessageRole.Assistant; } }
        public override List<MessageContent> Content { get; }
        public List<ChatToolCall> ToolCalls { get; set; } = new List<ChatToolCall>();

        public AssistantChatMessage(string content)
        {
            Content = new List<MessageContent> { new MessageContent(content) };
        }

        public AssistantChatMessage(ChatCompletion completion)
        {
            Content = new List<MessageContent> { new MessageContent(completion.Content ?? "") };
            ToolCalls = completion.ToolCalls ?? new List<ChatToolCall>();
        }
    }

    // 工具消息
    public class ToolChatMessage : ChatMessage
    {
        public override MessageRole Role { get { return MessageRole.Tool; } }
        public override List<MessageContent> Content { get; }
        public string ToolCallId { get; set; }

        public ToolChatMessage(string toolCallId, string content)
        {
            ToolCallId = toolCallId;
            Content = new List<MessageContent> { new MessageContent(content) };
        }
    }

    // 工具定义
    public class ChatTool
    {
        public string Type { get; set; } = "function";
        public FunctionDefinition Function { get; set; }

        public static ChatTool CreateFunctionTool(string functionName, string functionDescription, BinaryData functionParameters)
        {
            var parametersJson = Encoding.UTF8.GetString(functionParameters.ToArray());
            var parameters = JObject.Parse(parametersJson);

            return new ChatTool
            {
                Function = new FunctionDefinition
                {
                    Name = functionName,
                    Description = functionDescription,
                    Parameters = parameters
                }
            };
        }
    }

    // 函数定义
    public class FunctionDefinition
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public JObject Parameters { get; set; }
    }

    // 聊天完成选项
    public class ChatCompletionOptions
    {
        public List<ChatTool> Tools { get; set; } = new List<ChatTool>();
        public bool Stream { get; set; } = false;
        public double Temperature { get; set; } = 1.0;
        public int MaxTokens { get; set; } = 1000;
    }

    // 聊天完成结果
    public class ChatCompletion
    {
        public string Content { get; set; }
        public ChatFinishReason FinishReason { get; set; }
        public List<ChatToolCall> ToolCalls { get; set; } = new List<ChatToolCall>();
    }

    // 流式聊天完成结果
    public class StreamingChatCompletionUpdate
    {
        public string Content { get; set; }
        public ChatFinishReason FinishReason { get; set; }
        public List<ChatToolCall> ToolCalls { get; set; } = new List<ChatToolCall>();
        public bool IsFinished { get; set; }
    }

    // BinaryData 替代类 (C# 7.3 兼容)
    public class BinaryData
    {
        private byte[] _data;

        public BinaryData(byte[] data)
        {
            _data = data;
        }

        public byte[] ToArray()
        {
            return _data;
        }

        public static BinaryData FromBytes(byte[] data)
        {
            return new BinaryData(data);
        }

        public static BinaryData FromBytes(string data)
        {
            return new BinaryData(Encoding.UTF8.GetBytes(data));
        }
    }

    // DashScope API 请求模型
    internal class DashScopeRequest
    {
        [JsonProperty("model")]
        public string Model { get; set; }

        [JsonProperty("messages")]
        public List<DashScopeMessage> Messages { get; set; }

        [JsonProperty("tools")]
        public List<DashScopeTool> Tools { get; set; }

        [JsonProperty("stream")]
        public bool Stream { get; set; }

        [JsonProperty("temperature")]
        public double Temperature { get; set; }

        [JsonProperty("max_tokens")]
        public int MaxTokens { get; set; }
    }

    internal class DashScopeMessage
    {
        [JsonProperty("role")]
        public string Role { get; set; }

        [JsonProperty("content")]
        public string Content { get; set; }

        [JsonProperty("tool_calls")]
        public List<DashScopeToolCall> ToolCalls { get; set; }

        [JsonProperty("tool_call_id")]
        public string ToolCallId { get; set; }
    }

    internal class DashScopeToolCall
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("function")]
        public DashScopeFunction Function { get; set; }
    }

    internal class DashScopeFunction
    {
        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("arguments")]
        public string Arguments { get; set; }
    }

    internal class DashScopeTool
    {
        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("function")]
        public DashScopeFunctionDef Function { get; set; }
    }

    internal class DashScopeFunctionDef
    {
        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("description")]
        public string Description { get; set; }

        [JsonProperty("parameters")]
        public JObject Parameters { get; set; }
    }

    // DashScope API 响应模型
    internal class DashScopeResponse
    {
        [JsonProperty("choices")]
        public List<DashScopeChoice> Choices { get; set; }

        [JsonProperty("usage")]
        public DashScopeUsage Usage { get; set; }
    }

    internal class DashScopeChoice
    {
        [JsonProperty("message")]
        public DashScopeMessage Message { get; set; }

        [JsonProperty("finish_reason")]
        public string FinishReason { get; set; }

        [JsonProperty("delta")]
        public DashScopeMessage Delta { get; set; }
    }

    internal class DashScopeUsage
    {
        [JsonProperty("prompt_tokens")]
        public int PromptTokens { get; set; }

        [JsonProperty("completion_tokens")]
        public int CompletionTokens { get; set; }

        [JsonProperty("total_tokens")]
        public int TotalTokens { get; set; }
    }

    // 流式响应处理委托
    public delegate void StreamingResponseHandler(StreamingChatCompletionUpdate update);

    // Qwen 客户端
    public class QwenClient : IDisposable
    {
        private readonly string _model;
        private readonly string _apiKey;
        private readonly string _baseUrl;

        public QwenClient(string model, string apiKey, string baseUrl = "https://dashscope.aliyuncs.com")
        {
            _model = model;
            _apiKey = apiKey;
            _baseUrl = baseUrl.TrimEnd('/');
        }

        // 非流式聊天完成
        public async Task<ChatCompletion> CompleteChatAsync(List<ChatMessage> messages, ChatCompletionOptions options = null)
        {
            if (options == null)
                options = new ChatCompletionOptions();

            var request = BuildRequest(messages, options);
            var jsonRequest = JsonConvert.SerializeObject(request);

            System.Diagnostics.Debug.WriteLine($"发送请求到:{_baseUrl}/compatible-mode/v1/chat/completions");
            System.Diagnostics.Debug.WriteLine($"请求内容: {jsonRequest}");

            try
            {
                string result = await SendPostRequestAsync(_baseUrl + "/compatible-mode/v1/chat/completions", jsonRequest, _apiKey);
                System.Diagnostics.Debug.WriteLine($"收到响应: {result}");
                
                var dashScopeResponse = JsonConvert.DeserializeObject<DashScopeResponse>(result);
                return ConvertToChatCompletion(dashScopeResponse.Choices[0]);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"请求异常: {ex.Message}");
                throw;
            }
        }

        private async Task<string> SendPostRequestAsync(string url, string jsonContent, string apiKey)
        {
            using (var content = new StringContent(jsonContent, Encoding.UTF8, "application/json"))
            {
                // 创建新的HttpClient避免线程问题
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(30);
                    
                    // 设置请求头
                    client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                    System.Diagnostics.Debug.WriteLine($"完整请求URL: {url}");
                    System.Diagnostics.Debug.WriteLine($"请求头: Authorization=Bearer {apiKey.Substring(0, 10)}...");

                    // 发送请求并获取响应
                    HttpResponseMessage response = await client.PostAsync(url, content);

                    System.Diagnostics.Debug.WriteLine($"响应状态码: {response.StatusCode}");
                    System.Diagnostics.Debug.WriteLine($"响应头: {response.Headers}");

                    // 处理响应
                    if (response.IsSuccessStatusCode)
                    {
                        return await response.Content.ReadAsStringAsync();
                    }
                    else
                    {
                        string errorContent = await response.Content.ReadAsStringAsync();
                        System.Diagnostics.Debug.WriteLine($"错误内容: {errorContent}");
                        return $"请求失败: {response.StatusCode} - {errorContent}";
                    }
                }
            }
        }

        // 同步版本
        public ChatCompletion CompleteChat(List<ChatMessage> messages, ChatCompletionOptions options = null)
        {
            return CompleteChatAsync(messages, options).GetAwaiter().GetResult();
        }

        // 流式聊天完成 - 使用回调模式替代异步枚举
        public async Task CompleteChatStreamingAsync(List<ChatMessage> messages, StreamingResponseHandler onUpdate, ChatCompletionOptions options = null)
        {
            if (options == null)
                options = new ChatCompletionOptions();
            options.Stream = true;

            var request = BuildRequest(messages, options);
            var jsonRequest = JsonConvert.SerializeObject(request);

            using (var content = new StringContent(jsonRequest, Encoding.UTF8, "application/json"))
            {
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(60); // 流式请求需要更长时间
                    
                    // 设置请求头
                    client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", _apiKey);
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                    var response = await client.PostAsync(_baseUrl + "/compatible-mode/v1/chat/completions", content);
                    response.EnsureSuccessStatusCode();

                    var stream = await response.Content.ReadAsStreamAsync();
                    try
                    {
                        var reader = new StreamReader(stream);
                        try
                        {
                            string line;
                            while ((line = await reader.ReadLineAsync()) != null)
                            {
                                if (string.IsNullOrWhiteSpace(line) || !line.StartsWith("data: "))
                                    continue;

                                var data = line.Substring(6).Trim();
                                if (data == "[DONE]")
                                    break;

                                try
                                {
                                    var streamResponse = JsonConvert.DeserializeObject<DashScopeResponse>(data);
                                    if (streamResponse?.Choices?.Count > 0)
                                    {
                                        var update = ConvertToStreamingUpdate(streamResponse.Choices[0]);
                                        onUpdate?.Invoke(update);
                                    }
                                }
                                catch (JsonException)
                                {
                                    // 忽略解析错误
                                }
                            }
                        }
                        finally
                        {
                            reader.Dispose();
                        }
                    }
                    finally
                    {
                        stream.Dispose();
                    }
                }
            }
        }

        // 同步流式版本
        public void CompleteChatStreaming(List<ChatMessage> messages, StreamingResponseHandler onUpdate, ChatCompletionOptions options = null)
        {
            CompleteChatStreamingAsync(messages, onUpdate, options).GetAwaiter().GetResult();
        }

        private object BuildRequest(List<ChatMessage> messages, ChatCompletionOptions options)
        {
            // 创建基础请求对象
            var request = new Dictionary<string, object>
            {
                ["model"] = _model,
                ["messages"] = messages.Select(msg => 
                {
                    var msgDict = new Dictionary<string, object>
                    {
                        ["role"] = msg.ToJsonRole(),
                        ["content"] = msg.Content.FirstOrDefault()?.Text ?? ""
                    };
                    
                    // 只在有工具调用时才添加 tool_calls 字段
                    var assistantMsg = msg as AssistantChatMessage;
                    if (assistantMsg != null && assistantMsg.ToolCalls.Count > 0)
                    {
                        msgDict["tool_calls"] = assistantMsg.ToolCalls.Select(tc => new
                        {
                            id = tc.Id,
                            type = tc.Type,
                            function = new
                            {
                                name = tc.FunctionName,
                                arguments = tc.FunctionArguments
                            }
                        }).ToList();
                    }
                    
                    // 只在是工具消息时才添加 tool_call_id 字段
                    var toolMsg = msg as ToolChatMessage;
                    if (toolMsg != null)
                    {
                        msgDict["tool_call_id"] = toolMsg.ToolCallId;
                    }
                    
                    return msgDict;
                }).ToList()
            };

            // 只在流式模式时添加 stream 字段
            if (options.Stream)
            {
                request["stream"] = true;
            }

            // 只在有工具时添加 tools 字段
            if (options.Tools != null && options.Tools.Count > 0)
            {
                request["tools"] = options.Tools.Select(tool => new
                {
                    type = tool.Type,
                    function = new
                    {
                        name = tool.Function.Name,
                        description = tool.Function.Description,
                        parameters = tool.Function.Parameters
                    }
                }).ToList();
            }

            return request;
        }

        private ChatCompletion ConvertToChatCompletion(DashScopeChoice choice, bool isStreaming = false)
        {
            var message = isStreaming ? choice.Delta : choice.Message;

            var completion = new ChatCompletion
            {
                Content = message?.Content ?? "",
                FinishReason = ConvertFinishReason(choice.FinishReason)
            };

            if (message?.ToolCalls?.Count > 0)
            {
                completion.ToolCalls = message.ToolCalls.Select(tc => new ChatToolCall
                {
                    Id = tc.Id,
                    Type = tc.Type,
                    FunctionName = tc.Function.Name,
                    FunctionArguments = tc.Function.Arguments
                }).ToList();
            }

            return completion;
        }

        private StreamingChatCompletionUpdate ConvertToStreamingUpdate(DashScopeChoice choice)
        {
            var message = choice.Delta;

            var update = new StreamingChatCompletionUpdate
            {
                Content = message?.Content ?? "",
                FinishReason = ConvertFinishReason(choice.FinishReason),
                IsFinished = !string.IsNullOrEmpty(choice.FinishReason)
            };

            if (message?.ToolCalls?.Count > 0)
            {
                update.ToolCalls = message.ToolCalls.Select(tc => new ChatToolCall
                {
                    Id = tc.Id,
                    Type = tc.Type,
                    FunctionName = tc.Function.Name,
                    FunctionArguments = tc.Function.Arguments
                }).ToList();
            }

            return update;
        }

        private ChatFinishReason ConvertFinishReason(string finishReason)
        {
            if (finishReason == "stop")
                return ChatFinishReason.Stop;
            if (finishReason == "tool_calls")
                return ChatFinishReason.ToolCalls;
            if (finishReason == "length")
                return ChatFinishReason.Length;
            if (finishReason == "content_filter")
                return ChatFinishReason.ContentFilter;
            return ChatFinishReason.Stop;
        }

        public void Dispose()
        {
            // 不再需要处理 HttpClient，因为每次请求都会创建新的实例并自动释放
        }
    }

    // 兼容性包装类，用于替换原来的 ChatClient
    public class ChatClient : IDisposable
    {
        private readonly QwenClient _qwenClient;

        public ChatClient(string model, string apiKey, string baseUrl = "https://dashscope.aliyuncs.com")
        {
            _qwenClient = new QwenClient(model, apiKey, baseUrl);
        }

        public ChatCompletion CompleteChat(List<ChatMessage> messages, ChatCompletionOptions options = null)
        {
            return _qwenClient.CompleteChat(messages, options);
        }

        public async Task<ChatCompletion> CompleteChatAsync(List<ChatMessage> messages, ChatCompletionOptions options = null)
        {
            return await _qwenClient.CompleteChatAsync(messages, options);
        }

        public async Task CompleteChatStreamingAsync(List<ChatMessage> messages, StreamingResponseHandler onUpdate, ChatCompletionOptions options = null)
        {
            await _qwenClient.CompleteChatStreamingAsync(messages, onUpdate, options);
        }

        public void CompleteChatStreaming(List<ChatMessage> messages, StreamingResponseHandler onUpdate, ChatCompletionOptions options = null)
        {
            _qwenClient.CompleteChatStreaming(messages, onUpdate, options);
        }

        public void Dispose()
        {
            _qwenClient?.Dispose();
        }
    }
}