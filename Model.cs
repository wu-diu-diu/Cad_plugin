using OpenAI;

using OpenAI.Chat;
using System;
using System.ClientModel;

namespace Model
{
    class Program
    {
        static void Main(string[] args)
        {
            ChatClient client = new ChatClient(model: "gpt-4o",
                credential: new ApiKeyCredential(Environment.GetEnvironmentVariable("OPENAI_API_KEY")),
                options: new OpenAIClientOptions()
                {
                    Endpoint = new Uri("https://api.nuwaapi.com/v1")
                });

            ChatCompletion completion = client.CompleteChat("hello!");

            Console.WriteLine($"[ASSISTANT]: {completion.Content[0].Text}");
        }
    }
}