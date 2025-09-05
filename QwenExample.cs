using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using QwenSDK;
using Newtonsoft.Json.Linq;

namespace QwenExample
{
    class Program
    {
        // 替换为你的实际API Key
        private const string API_KEY = "your-dashscope-api-key";

        static async Task Main(string[] args)
        {
            var client = new QwenClient(API_KEY);

            Console.WriteLine("=== Qwen C# SDK 演示 ===\n");

            try
            {
                // 演示1: 非流式输出
                await DemoNonStreaming(client);

                Console.WriteLine("\n" + new string('=', 50) + "\n");

                // 演示2: 流式输出
                await DemoStreaming(client);

                Console.WriteLine("\n" + new string('=', 50) + "\n");

                // 演示3: 工具调用
                await DemoToolCalling(client);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"发生错误: {ex.Message}");
            }
            finally
            {
                client.Dispose();
            }

            Console.WriteLine("\n按任意键退出...");
            Console.ReadKey();
        }

        /// <summary>
        /// 演示非流式输出
        /// </summary>
        static async Task DemoNonStreaming(QwenClient client)
        {
            Console.WriteLine("【演示1: 非流式输出】");

            // 方法1: 使用便捷方法
            Console.WriteLine("使用便捷方法:");
            var simpleResponse = await client.ChatAsync(
                "请用一句话介绍杭州这座城市",
                "你是一个旅游助手"
            );
            Console.WriteLine($"回答: {simpleResponse}");

            Console.WriteLine("\n使用完整API:");
            // 方法2: 使用完整API
            var messages = new List<QwenMessage>
            {
                new QwenMessage { Role = "system", Content = "你是一个历史学家" },
                new QwenMessage { Role = "user", Content = "简单介绍一下中国古代四大发明" }
            };

            var request = new QwenRequest
            {
                Model = "qwen-plus",
                Messages = messages,
                Temperature = 0.7,
                MaxTokens = 1000
            };

            var response = await client.ChatCompletionAsync(request);
            if (response.Choices != null && response.Choices.Count > 0)
            {
                Console.WriteLine($"回答: {response.Choices[0].Message.Content}");
                Console.WriteLine($"使用tokens: {response.Usage?.TotalTokens}");
            }
        }

        /// <summary>
        /// 演示流式输出
        /// </summary>
        static async Task DemoStreaming(QwenClient client)
        {
            Console.WriteLine("【演示2: 流式输出】");

            // 方法1: 使用便捷方法
            Console.WriteLine("使用便捷流式方法:");
            Console.Write("回答: ");

            await client.ChatStreamAsync(
                "请详细介绍一下人工智能的发展历程",
                "你是一个AI专家",
                chunk => Console.Write(chunk) // 实时输出每个chunk
            );

            Console.WriteLine("\n\n使用完整流式API:");
            // 方法2: 使用完整流式API
            var streamMessages = new List<QwenMessage>
            {
                new QwenMessage { Role = "system", Content = "你是一个科学家" },
                new QwenMessage { Role = "user", Content = "解释一下量子计算的基本原理" }
            };

            var streamRequest = new QwenRequest
            {
                Model = "qwen-plus",
                Messages = streamMessages,
                Temperature = 0.8
            };

            Console.Write("回答: ");

            // 订阅流式事件
            client.StreamChunkReceived += (sender, e) =>
            {
                if (e.Chunk?.Choices != null && e.Chunk.Choices.Count > 0)
                {
                    var delta = e.Chunk.Choices[0].Delta;
                    if (!string.IsNullOrEmpty(delta?.Content))
                    {
                        Console.Write(delta.Content);
                    }
                }
            };

            client.StreamCompleted += (sender, e) =>
            {
                Console.WriteLine("\n[流式输出完成]");
            };

            client.StreamError += (sender, e) =>
            {
                Console.WriteLine($"\n[流式输出错误: {e.Message}]");
            };

            await client.ChatCompletionStreamAsync(streamRequest);
        }

        /// <summary>
        /// 演示工具调用
        /// </summary>
        static async Task DemoToolCalling(QwenClient client)
        {
            Console.WriteLine("【演示3: 工具调用】");

            // 定义可用工具
            var tools = new List<QwenTool>
            {
                QwenToolBuilder.CreateWeatherTool(),
                QwenToolBuilder.CreateTimeTool(),
                QwenToolBuilder.CreateFunction(
                    "calculate",
                    "执行基本的数学计算",
                    new
                    {
                        type = "object",
                        properties = new
                        {
                            expression = new
                            {
                                type = "string",
                                description = "数学表达式，比如 '2+3*4'"
                            }
                        },
                        required = new[] { "expression" }
                    }
                )
            };

            var messages = new List<QwenMessage>
            {
                new QwenMessage
                {
                    Role = "system",
                    Content = "你是一个智能助手，可以查询天气、获取时间和进行计算。当用户询问相关信息时，请使用对应的工具。"
                },
                new QwenMessage
                {
                    Role = "user",
                    Content = "现在几点了？另外杭州今天天气怎么样？"
                }
            };

            // 发送带工具的请求
            var response = await client.ChatCompletionWithToolsAsync(messages, tools);

            if (response.Choices != null && response.Choices.Count > 0)
            {
                var choice = response.Choices[0];

                if (choice.Message.ToolCalls != null && choice.Message.ToolCalls.Count > 0)
                {
                    Console.WriteLine("AI请求调用工具:");

                    foreach (var toolCall in choice.Message.ToolCalls)
                    {
                        Console.WriteLine($"工具: {toolCall.Function.Name}");
                        Console.WriteLine($"参数: {toolCall.Function.Arguments}");

                        // 模拟执行工具调用
                        string toolResult = await SimulateToolCall(toolCall);

                        // 将工具结果添加到消息历史
                        messages.Add(choice.Message); // AI的工具调用消息
                        messages.Add(new QwenMessage
                        {
                            Role = "tool",
                            ToolCallId = toolCall.Id,
                            Name = toolCall.Function.Name,
                            Content = toolResult
                        });
                    }

                    // 发送包含工具结果的后续请求
                    var followUpResponse = await client.ChatCompletionWithToolsAsync(messages, tools);

                    if (followUpResponse.Choices != null && followUpResponse.Choices.Count > 0)
                    {
                        Console.WriteLine("\nAI最终回答:");
                        Console.WriteLine(followUpResponse.Choices[0].Message.Content);
                    }
                }
                else
                {
                    Console.WriteLine("AI直接回答（未使用工具）:");
                    Console.WriteLine(choice.Message.Content);
                }
            }

            Console.WriteLine("\n--- 演示工具调用的流式输出 ---");

            // 演示流式工具调用
            var streamMessages = new List<QwenMessage>
            {
                new QwenMessage
                {
                    Role = "user",
                    Content = "帮我计算一下 (15 + 25) * 2 等于多少？"
                }
            };

            client.StreamChunkReceived += (sender, e) =>
            {
                if (e.Chunk?.Choices != null && e.Chunk.Choices.Count > 0)
                {
                    var delta = e.Chunk.Choices[0].Delta;
                    if (!string.IsNullOrEmpty(delta?.Content))
                    {
                        Console.Write(delta.Content);
                    }

                    // 检查是否有工具调用
                    if (delta?.ToolCalls != null)
                    {
                        Console.WriteLine("\n[检测到工具调用请求]");
                        foreach (var toolCall in delta.ToolCalls)
                        {
                            Console.WriteLine($"工具: {toolCall.Function?.Name}");
                            Console.WriteLine($"参数: {toolCall.Function?.Arguments}");
                        }
                    }
                }
            };

            Console.Write("流式回答: ");
            await client.ChatCompletionWithToolsStreamAsync(streamMessages, tools);
        }

        /// <summary>
        /// 模拟工具调用执行
        /// </summary>
        static async Task<string> SimulateToolCall(ToolCall toolCall)
        {
            await Task.Delay(100); // 模拟网络延迟

            switch (toolCall.Function.Name)
            {
                case "get_current_time":
                    return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                case "get_current_weather":
                    try
                    {
                        var args = JObject.Parse(toolCall.Function.Arguments);
                        var location = args["location"]?.ToString();
                        // 这里应该调用真实的天气API
                        return $"{location}今天天气晴朗，温度25°C，湿度60%";
                    }
                    catch
                    {
                        return "无法获取天气信息";
                    }

                case "calculate":
                    try
                    {
                        var args = JObject.Parse(toolCall.Function.Arguments);
                        var expression = args["expression"]?.ToString();
                        // 这里应该使用安全的表达式计算器
                        // 为了演示，我们只处理简单的情况
                        if (expression == "(15 + 25) * 2")
                        {
                            return "80";
                        }
                        return $"计算结果: {expression} = [需要实现计算逻辑]";
                    }
                    catch
                    {
                        return "计算失败";
                    }

                default:
                    return $"未知工具: {toolCall.Function.Name}";
            }
        }
    }
}