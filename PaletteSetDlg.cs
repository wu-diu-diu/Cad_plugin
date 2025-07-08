using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using DeepSeek.Sdk;
using Markdig;  // 支持Markdown语法
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms; 
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using Autodesk.AutoCAD.Runtime;
using CADApplication = Autodesk.AutoCAD.ApplicationServices.Application;

namespace CoDesignStudy.Cad.PlugIn
{
    public partial class PaletteSetDlg : UserControl
    {
        private Panel conversationPanel;
        private TextBox inputTextBox;
        private Button sendButton;

        public PaletteSetDlg()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Size = new Size(800, 600);
            this.MinimumSize = new Size(400, 300);

            // 对话区域（使用Panel以支持自定义消息气泡布局）
            conversationPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.White,
                FlowDirection = System.Windows.Forms.FlowDirection.TopDown,
                WrapContents = false
            };
            conversationPanel.Resize += ConversationPanel_Resize;

            // 输入区域
            var inputPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                BackColor = Color.WhiteSmoke,
                Padding = new Padding(10, 5, 10, 5)
            };

            inputTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true
            };
            inputTextBox.KeyDown += InputTextBox_KeyDown;
            inputPanel.Controls.Add(inputTextBox);

            sendButton = new Button
            {
                Dock = DockStyle.Right,
                Text = "发送",
                Width = 80
            };
            sendButton.Click += SendButton_Click;
            inputPanel.Controls.Add(sendButton);

            this.Controls.Add(conversationPanel);
            this.Controls.Add(inputPanel);
        }

        private async void SendButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(inputTextBox.Text))
                return;

            string userMessage = inputTextBox.Text.Trim();
            inputTextBox.Clear();

            await AppendMessageAsync("用户", userMessage);

            sendButton.Enabled = false;

            try
            {
                Action<string> updateFunc = null;
                // 等待 AI 控件初始化并获得更新函数
                var aiContentControl = await AppendMessageAsync("AI", "", true, setter => updateFunc = setter);
                await GetAIResponse(userMessage, updateFunc);
            }
            finally
            {
                sendButton.Enabled = true;
            }
        }

        public async Task<Control> AppendMessageAsync(string sender, string message, bool isStreaming = false, Action<Action<string>> setUpdateContent = null)
        {
            if (string.IsNullOrWhiteSpace(message) && !isStreaming) return null;

            Control contentControl;

            if (sender == "AI")
            {
                var webView = new WebView2
                {
                    Width = conversationPanel.Width - 150,
                    Margin = new Padding(0),
                    Height = 1,
                    DefaultBackgroundColor = Color.White
                };

                await webView.EnsureCoreWebView2Async();

                if (webView.CoreWebView2 != null)
                {
                    webView.CoreWebView2.WebMessageReceived += (s, a) =>
                    {
                        if (int.TryParse(a.WebMessageAsJson, out int height))
                        {
                            Action update = () =>
                            {
                                webView.Height = Math.Max(height, 1);
                                // ⭐ 关键：更新高度后再滚动
                                conversationPanel.ScrollControlIntoView(webView);
                            };

                            if (webView.InvokeRequired)
                                webView.Invoke(update);
                            else
                                update();
                        }
                    };

                    string html = Markdig.Markdown.ToHtml(message ?? "");
                    string doc = $"<html><head><script>window.MathJax={{tex:{{inlineMath:[['$','$'],['\\\\(','\\\\)']],displayMath:[['$$','$$'],['\\\\[','\\\\]']]}},svg:{{fontCache:'global'}}}};</script><script src='https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js'></script></head><body id='body' style='font-family:微软雅黑;font-size:10pt;background-color:#FFFFFF;margin:0;padding:0;'>{html}<script>function setContent(html){{const body=document.getElementById('body');body.innerHTML=html;if(window.MathJax){{MathJax.typesetPromise([body]).then(()=>{{setTimeout(()=>{{const height=Math.max(body.scrollHeight,document.documentElement.scrollHeight);window.chrome.webview.postMessage(height);}},100);}});}}else{{setTimeout(()=>{{const height=Math.max(body.scrollHeight,document.documentElement.scrollHeight);window.chrome.webview.postMessage(height);}},100);}}const resizeObserver=new ResizeObserver(()=>{{const newHeight=Math.max(body.scrollHeight,document.documentElement.scrollHeight);window.chrome.webview.postMessage(newHeight);}});resizeObserver.observe(body);}}if(window.MathJax)MathJax.typesetPromise();const initialHeight=Math.max(document.body.scrollHeight,document.documentElement.scrollHeight);window.chrome.webview.postMessage(initialHeight);</script></body></html>";
                    webView.CoreWebView2.NavigateToString(doc);

                    if (setUpdateContent != null)
                    {
                        Action<string> updateContent = (newContent) =>
                        {
                            string htmlUpdate = Markdig.Markdown.ToHtml(newContent).Replace("`", "\\`");

                            if (webView.InvokeRequired)
                            {
                                webView.Invoke(new Action(() =>
                                {
                                    webView.CoreWebView2.ExecuteScriptAsync($"setContent(`{htmlUpdate}`);");
                                    conversationPanel.ScrollControlIntoView(webView);
                                    // ❌ 不在这里滚动
                                }));
                            }
                            else
                            {
                                webView.CoreWebView2.ExecuteScriptAsync($"setContent(`{htmlUpdate}`);");
                                conversationPanel.ScrollControlIntoView(webView);
                                // ❌ 不在这里滚动
                            }
                        };

                        setUpdateContent(updateContent);
                    }
                }

                contentControl = webView;
            }
            else
            {
                contentControl = new Label
                {
                    Text = message,
                    AutoSize = true,
                    Font = new System.Drawing.Font("微软雅黑", 10),
                    MaximumSize = new Size(conversationPanel.Width - 150, 0),
                    Padding = new Padding(10),
                    BackColor = Color.LightBlue,
                    Margin = new Padding(5)
                };
            }

            // 头像
            var avatar = new PictureBox
            {
                Size = new Size(40, 40),
                SizeMode = PictureBoxSizeMode.Zoom,
                Margin = new Padding(5),
                Image = sender == "用户"
                    ? System.Drawing.Image.FromFile("userAvatar.png")
                    : System.Drawing.Image.FromFile("aiAvatar.png")
            };

            var horizontalPanel = new FlowLayoutPanel
            {
                AutoSize = true,
                FlowDirection = sender == "用户" ? System.Windows.Forms.FlowDirection.RightToLeft : System.Windows.Forms.FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };

            horizontalPanel.Controls.Add(avatar);
            horizontalPanel.Controls.Add(contentControl);
            // TODO: 鼠标滚动只能在侧边栏的最右侧实现，可能是因为ai回复消息的气泡的控件不支持鼠标滚动
            var container = new FlowLayoutPanel
            {
                AutoSize = true,                       // ✅ 必须，确保高度自适应
                Anchor = AnchorStyles.Left,           // ✅ 防止错位
                Margin = new Padding(0),              // ✅ 清除默认边距
                Padding = new Padding(0),
                FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                WrapContents = false
            };

            if (sender == "用户")
            {
                container.Controls.Add(new Label() { Width = conversationPanel.Width - horizontalPanel.PreferredSize.Width - 30, AutoSize = false });
                container.Controls.Add(horizontalPanel);
            }
            else
            {
                container.Controls.Add(horizontalPanel);
                container.Controls.Add(new Label() { Width = conversationPanel.Width - horizontalPanel.PreferredSize.Width - 30, AutoSize = false });
            }

            conversationPanel.Controls.Add(container);
            //conversationPanel.Controls.SetChildIndex(container, 0);

            // 滚动到底
            conversationPanel.Controls.Add(new Panel() { Height = 10, Dock = DockStyle.Top });
            conversationPanel.ScrollControlIntoView(container);

            return contentControl;
        }

        public Task<Control> AppendMessageSync(string sender, string message)
        {
            if (string.IsNullOrWhiteSpace(message)) return Task.FromResult<Control>(null);

            Control contentControl;

            // AI 使用 WebView 渲染 markdown，用户使用 Label
            if (sender == "AI")
            {
                var webView = new WebView2
                {
                    Width = conversationPanel.Width - 150,
                    Margin = new Padding(0),
                    Height = 1,
                    DefaultBackgroundColor = Color.White
                };

                webView.EnsureCoreWebView2Async().Wait();

                string html = Markdig.Markdown.ToHtml(message);
                string doc = $@"
<html><head>
<script src='https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js'></script>
</head>
<body style='font-family:微软雅黑;font-size:10pt;background-color:#FFFFFF;margin:0;padding:0;'>
{html}
</body></html>";

                webView.CoreWebView2.NavigateToString(doc);
                contentControl = webView;
            }
            else
            {
                contentControl = new Label
                {
                    Text = message,
                    AutoSize = true,
                    Font = new System.Drawing.Font("微软雅黑", 10),
                    MaximumSize = new Size(conversationPanel.Width - 150, 0),
                    Padding = new Padding(10),
                    BackColor = Color.LightBlue,
                    Margin = new Padding(5)
                };
            }

            // 头像
            var avatar = new PictureBox
            {
                Size = new Size(40, 40),
                SizeMode = PictureBoxSizeMode.Zoom,
                Margin = new Padding(5),
                Image = sender == "用户"
                    ? System.Drawing.Image.FromFile("userAvatar.png")
                    : System.Drawing.Image.FromFile("aiAvatar.png")
            };

            var horizontalPanel = new FlowLayoutPanel
            {
                AutoSize = true,
                FlowDirection = sender == "用户" ? System.Windows.Forms.FlowDirection.RightToLeft : System.Windows.Forms.FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };

            horizontalPanel.Controls.Add(avatar);
            horizontalPanel.Controls.Add(contentControl);
            // TODO: 鼠标滚动只能在侧边栏的最右侧实现，可能是因为ai回复消息的气泡的控件不支持鼠标滚动
            var container = new FlowLayoutPanel
            {
                AutoSize = true,                       // ✅ 必须，确保高度自适应
                Anchor = AnchorStyles.Left,           // ✅ 防止错位
                Margin = new Padding(0),              // ✅ 清除默认边距
                Padding = new Padding(0),
                FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                WrapContents = false
            };

            if (sender == "用户")
            {
                container.Controls.Add(new Label()
                {
                    Width = conversationPanel.Width - horizontalPanel.PreferredSize.Width - 30,
                    AutoSize = false
                });
                container.Controls.Add(horizontalPanel);
            }
            else
            {
                container.Controls.Add(horizontalPanel);
                container.Controls.Add(new Label()
                {
                    Width = conversationPanel.Width - horizontalPanel.PreferredSize.Width - 30,
                    AutoSize = false
                });
            }

            conversationPanel.Controls.Add(container);
            conversationPanel.Controls.Add(new Panel() { Height = 10, Dock = DockStyle.Top });
            conversationPanel.ScrollControlIntoView(container);

            return Task.FromResult<Control>(contentControl);
        }


        public async Task<string> GetAIResponse(string userMessage, Action<string> updateContent)
        {
            string apikey = "sk-8812dc6bd29845c897813c3cfeb83a34";

            var ds = new DeepSeek.Sdk.DeepSeek(apikey);

            var resultMsg = new StringBuilder();

            var chatReq = new ChatRequest
            {
                Messages = new List<ChatRequest.MessagesType>
        {
            new ChatRequest.MessagesType
            {
                Role = ChatRequest.RoleEnum.System,
                Content = "你的名字是\"电气设计助手\"，是一个电气设计领域的AutoCAD助手，请回答有关电气设计领域的CAD操作的问题或者相关的电气知识"
            },
            new ChatRequest.MessagesType
            {
                Role = ChatRequest.RoleEnum.Assistant,
                Content = ""
            },
            new ChatRequest.MessagesType
            {
                Role = ChatRequest.RoleEnum.User,
                Content = userMessage
            }
        },
                Model = ChatRequest.ModelEnum.DeepseekChat,
                Stream = true
            };

            var tcs = new TaskCompletionSource<string>();

            await ds.ChatStream(chatReq,
                openedCallBack: (state) => { },
                closedCallBack: (state) => { tcs.TrySetResult(resultMsg.ToString()); },
                msgCallback: async (res) =>
                {
                    string msg = res.Choices.FirstOrDefault()?.Delta?.Content;
                    if (!string.IsNullOrEmpty(msg))
                    {
                        resultMsg.Append(msg);
                        updateContent?.Invoke(resultMsg.ToString());
                    }
                },
                errorCallback: (ex) =>
                {
                    tcs.TrySetException(new Exception(ex));
                });

            return await tcs.Task;
        }
        public static async Task<string> CallLLMAsync(string prompt)
        {
            string apikey = "sk-8812dc6bd29845c897813c3cfeb83a34";

            var ds = new DeepSeek.Sdk.DeepSeek(apikey);

            var chatReq = new ChatRequest
            {
                Messages = new List<ChatRequest.MessagesType>
        {
            new ChatRequest.MessagesType
            {
                Role = ChatRequest.RoleEnum.System,
                Content = "你是一个Excel办公助手,仅回答办公相关的内容,其他无关内容忽略。"
            },
            new ChatRequest.MessagesType
            {
                Role = ChatRequest.RoleEnum.User,
                Content = prompt
            }
        },
                Model = ChatRequest.ModelEnum.DeepseekChat,
                Stream = false // ✅ 非流式调用
            };

            var chatRes = await ds.Chat(chatReq); // 非流式接口
            return chatRes.Choices.FirstOrDefault()?.Message?.Content ?? "";
        }


        private void InputTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && !e.Shift)
            {
                e.SuppressKeyPress = true;
                SendButton_Click(sender, e);
            }
        }

        // TODO: 实现聊天记录清除按键
        private void ClearButton_Click(object sender, EventArgs e)
        {
            conversationPanel.Controls.Clear();
        }

        private void ConversationPanel_Resize(object sender, EventArgs e)
        {
            // 重新调整 Label 最大宽度
            foreach (Control ctrl in conversationPanel.Controls)
            {
                if (ctrl is Panel panel)
                {
                    foreach (Control child in panel.Controls)
                    {
                        if (child is Label lbl)
                        {
                            lbl.MaximumSize = new Size(conversationPanel.Width - 100, 0);
                        }
                    }
                }
            }
        }
    }
}
