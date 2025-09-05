# Qwen C# SDK 使用指南

这是一个完整的Qwen API C#客户端实现，支持**非流式输出**、**流式输出**和**工具调用**三种功能。

## 🚀 功能特性

- ✅ **非流式聊天完成** - 一次性获取完整回答
- ✅ **流式聊天完成** - 实时接收回答内容
- ✅ **工具调用支持** - 支持函数调用和工具集成
- ✅ **完整的数据模型** - 类型安全的请求和响应处理
- ✅ **事件驱动架构** - 支持流式数据的事件处理
- ✅ **异常处理** - 完善的错误处理机制
- ✅ **便捷方法** - 提供简化的API调用方式

## 📦 安装依赖

```xml
<PackageReference Include="Newtonsoft.Json" Version="13.0.3" />
```

## 🔧 基本用法

### 1. 初始化客户端

```csharp
using QwenSDK;

// 替换为你的DashScope API Key
var client = new QwenClient("your-dashscope-api-key");
```

### 2. 非流式输出

```csharp
// 方法1: 便捷方法
string response = await client.ChatAsync(
    "你好，请介绍一下自己", 
    "你是一个AI助手"
);
Console.WriteLine(response);

// 方法2: 完整API
var messages = new List<QwenMessage>
{
    new QwenMessage { Role = "system", Content = "你是一个专家" },
    new QwenMessage { Role = "user", Content = "解释量子计算" }
};

var request = new QwenRequest
{
    Model = "qwen-plus",
    Messages = messages,
    Temperature = 0.7
};

var result = await client.ChatCompletionAsync(request);
Console.WriteLine(result.Choices[0].Message.Content);
```

### 3. 流式输出

```csharp
// 方法1: 便捷流式方法
await client.ChatStreamAsync(
    "写一首关于春天的诗",
    "你是一个诗人",
    chunk => Console.Write(chunk) // 实时输出
);

// 方法2: 事件驱动方式
client.StreamChunkReceived += (sender, e) =>
{
    if (e.Chunk?.Choices?[0]?.Delta?.Content != null)
    {
        Console.Write(e.Chunk.Choices[0].Delta.Content);
    }
};

client.StreamCompleted += (sender, e) =>
{
    Console.WriteLine("\n[完成]");
};

var streamRequest = new QwenRequest
{
    Model = "qwen-plus", 
    Messages = messages
};

await client.ChatCompletionStreamAsync(streamRequest);
```

### 4. 工具调用

```csharp
// 定义工具
var tools = new List<QwenTool>
{
    QwenToolBuilder.CreateWeatherTool(), // 内置天气工具
    QwenToolBuilder.CreateTimeTool(),    // 内置时间工具
    QwenToolBuilder.CreateFunction(      // 自定义工具
        "calculate",
        "执行数学计算",
        new
        {
            type = "object",
            properties = new
            {
                expression = new
                {
                    type = "string",
                    description = "数学表达式"
                }
            },
            required = new[] { "expression" }
        }
    )
};

// 发送工具调用请求
var messages = new List<QwenMessage>
{
    new QwenMessage { Role = "user", Content = "现在几点？杭州天气如何？" }
};

var response = await client.ChatCompletionWithToolsAsync(messages, tools);

// 处理工具调用
if (response.Choices[0].Message.ToolCalls != null)
{
    foreach (var toolCall in response.Choices[0].Message.ToolCalls)
    {
        Console.WriteLine($"调用工具: {toolCall.Function.Name}");
        Console.WriteLine($"参数: {toolCall.Function.Arguments}");
        
        // 执行工具并返回结果
        string toolResult = ExecuteTool(toolCall);
        
        // 添加工具结果到对话
        messages.Add(response.Choices[0].Message);
        messages.Add(new QwenMessage
        {
            Role = "tool",
            ToolCallId = toolCall.Id,
            Name = toolCall.Function.Name,
            Content = toolResult
        });
    }
    
    // 获取最终回答
    var finalResponse = await client.ChatCompletionWithToolsAsync(messages, tools);
    Console.WriteLine(finalResponse.Choices[0].Message.Content);
}
```

## 🛠️ 高级功能

### 自定义参数

```csharp
var request = new QwenRequest
{
    Model = "qwen-plus",
    Messages = messages,
    Temperature = 0.8,      // 创造性控制
    MaxTokens = 2000,       // 最大token数
    Tools = tools           // 工具列表
};
```

### 错误处理

```csharp
try
{
    var response = await client.ChatCompletionAsync(request);
}
catch (HttpRequestException ex)
{
    Console.WriteLine($"API请求失败: {ex.Message}");
}
catch (Exception ex)
{
    Console.WriteLine($"其他错误: {ex.Message}");
}
```

### 流式错误处理

```csharp
client.StreamError += (sender, e) =>
{
    Console.WriteLine($"流式处理错误: {e.Message}");
};
```

## 📝 数据模型说明

### QwenMessage
- `Role`: "system", "user", "assistant", "tool"
- `Content`: 消息内容
- `ToolCalls`: 工具调用信息（可选）
- `ToolCallId`: 工具调用ID（工具回复时使用）

### QwenTool
- `Type`: "function"
- `Function`: 函数定义信息

### QwenRequest
- `Model`: 模型名称（qwen-plus, qwen-turbo等）
- `Messages`: 消息列表
- `Tools`: 可用工具列表
- `Stream`: 是否流式输出
- `Temperature`: 随机性控制
- `MaxTokens`: 最大输出长度

## 🌟 最佳实践

1. **资源管理**: 使用完毕后调用 `client.Dispose()`
2. **API Key安全**: 不要在代码中硬编码API Key，使用环境变量或配置文件
3. **错误处理**: 总是包装API调用在try-catch块中
4. **流式处理**: 对于长文本生成，优先使用流式API
5. **工具调用**: 合理设计工具参数schema，确保AI能正确理解

## 📖 完整示例

参考 `QwenExample.cs` 文件查看完整的使用演示，包括：
- 非流式对话
- 流式对话
- 工具调用
- 错误处理

## 🔗 相关链接

- [DashScope API文档](https://help.aliyun.com/zh/dashscope/)
- [Qwen模型介绍](https://github.com/QwenLM/Qwen)