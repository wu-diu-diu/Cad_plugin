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
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static LaunchDarkly.Logging.LogCapture;

//using Autodesk.AutoCAD.Runtime;
using CADApplication = Autodesk.AutoCAD.ApplicationServices.Application;

namespace CoDesignStudy.Cad.PlugIn
{
    public partial class PaletteSetDlg : UserControl
    {
        private Panel conversationPanel;
        private TextBox inputTextBox;
        private Button sendButton;
        private static readonly int fontsize = 14;
        private static readonly string htmlsize = $"{fontsize}pt";

        // 自动滚动相关变量
        private bool isAutoScrollEnabled = true;
        private int scrollThreshold = 50; // 距离底部50px内认为在底部
        private DateTime lastUserScrollTime = DateTime.MinValue;
        private Control currentAIControl = null; // 当前正在更新的AI控件

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
            conversationPanel.Scroll += ConversationPanel_Scroll;
            conversationPanel.MouseWheel += ConversationPanel_MouseWheel;
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
        #region 滚动检测和控制

        /// <summary>
        /// 检测是否接近底部
        /// </summary>
        private bool IsNearBottom()
        {
            if (conversationPanel.Controls.Count == 0) return true;

            var verticalScroll = conversationPanel.VerticalScroll;
            int currentScrollPosition = verticalScroll.Value;
            int maxScrollPosition = verticalScroll.Maximum - conversationPanel.Height;

            return (maxScrollPosition - currentScrollPosition) <= scrollThreshold;
        }
        /// <summary>
        /// 滚动事件处理
        /// </summary>
        private void ConversationPanel_Scroll(object sender, ScrollEventArgs e)
        {
            // 记录用户手动滚动的时间
            if (e.Type == ScrollEventType.ThumbTrack || e.Type == ScrollEventType.ThumbPosition)
            {
                lastUserScrollTime = DateTime.Now;

                // 如果用户滚动到底部，重新启用自动滚动
                if (IsNearBottom())
                {
                    isAutoScrollEnabled = true;
                }
                else
                {
                    // 如果用户向上滚动，暂时禁用自动滚动
                    isAutoScrollEnabled = false;
                }
            }
        }
        /// <summary>
        /// 鼠标滚轮事件处理
        /// </summary>
        private void ConversationPanel_MouseWheel(object sender, MouseEventArgs e)
        {
            lastUserScrollTime = DateTime.Now;

            // 检查滚动方向和位置
            if (e.Delta < 0) // 向下滚动
            {
                if (IsNearBottom())
                {
                    isAutoScrollEnabled = true;
                }
            }
            else // 向上滚动
            {
                isAutoScrollEnabled = false;
            }
        }
        /// <summary>
        /// 智能自动滚动
        /// </summary>
        private void SmartAutoScroll(Control targetControl = null)
        {
            if (!isAutoScrollEnabled) return;

            // 如果用户刚刚手动滚动（1秒内），则不自动滚动
            if ((DateTime.Now - lastUserScrollTime).TotalMilliseconds < 1000)
            {
                return;
            }

            try
            {
                if (targetControl != null)
                {
                    // 滚动到指定控件
                    conversationPanel.ScrollControlIntoView(targetControl);
                }
                else
                {
                    // 滚动到底部
                    if (conversationPanel.Controls.Count > 0)
                    {
                        var lastControl = conversationPanel.Controls[conversationPanel.Controls.Count - 1];
                        conversationPanel.ScrollControlIntoView(lastControl);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Auto scroll error: {ex.Message}");
            }
        }

        /// <summary>
        /// 强制滚动到底部（新消息时使用）
        /// </summary>
        private void ScrollToBottom()
        {
            isAutoScrollEnabled = true;
            SmartAutoScroll();
        }

        #endregion

        #region 消息处理
        public async Task<string> SendAsync(string message, string instruction)
        {
            if (string.IsNullOrWhiteSpace(message)) return null;

            //await AppendMessageAsync("用户", instruction);


            string fullAIResponse = "";

            try
            {
                Action<string> updateFunc = null;
                var aiContentControl = await AppendMessageAsync("AI", "", true, setter => updateFunc = setter);
                fullAIResponse = await GetAIResponse(message, updateFunc);
            }
            catch (Exception ex)
            {
                System.IO.File.AppendAllText("C:\\ai_debug_log.txt", $"[SendAsync ERROR] {ex.Message}\n");
                throw;
            }
            finally
            {
                
            }
            return fullAIResponse;
        }
        private async void SendButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(inputTextBox.Text))
                return;

            string userMessage = inputTextBox.Text.Trim();
            inputTextBox.Clear();

            await AppendMessageAsync("用户", userMessage);
            ScrollToBottom();

            if (userMessage.EndsWith("材料清单"))
            {

                // 调用导出方法（注意：如果该方法不是线程安全的，可以考虑封装到 Task.Run）
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("EE ", true, false, false);
                await Task.Delay(2000);
                await AppendMessageAsync("AI", "已为您生成材料清单");
                return;
            }
            if (userMessage.StartsWith("修改"))
            {
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("delete_test ", true, false, false);
                var (roomType, coords, doorCoords) = InsertTracker.GetLastRoomDrawingInputs();
                userMessage = Prompt.GetLightingPrompt(roomType, coords, doorCoords, userMessage);
                string reply = await SendAsync(userMessage, "");
                var match = Regex.Match(reply, @"```json\s*([\s\S]+?)\s*```");

                string ModelReplyJson = match.Groups[1].Value;
                var commandInstance = new Command();
                commandInstance.InsertLightingFromModelReply(ModelReplyJson);
                return;
            }

            sendButton.Enabled = false;

            try
            {
                Action<string> updateFunc = null;
                // 等待 AI 控件初始化并获得更新函数
                var aiContentControl = await AppendMessageAsync("AI", "", true, setter => updateFunc = setter);
                // 设置当前AI控件
                currentAIControl = aiContentControl;
                await GetAIResponse(userMessage, updateFunc);
            }
            finally
            {
                sendButton.Enabled = true;
                currentAIControl = null;
            }
        }
        // 流式输出
        /// <param name="sender">消息发送者</param>
        /// <param name="message">消息内容</param>
        /// <param name="isStreaming">是否流式输出（单位 W）</param>
        /// <param name="setUpdateContent">回调函数，用于将AI内容逐步更新到控件中（单位 m）</param>
        /// <return return="Task<>是net中表示异步操作的类型，control是system.windows.forms.control类型的对象"> </return>
        public async Task<Control> AppendMessageAsync(string sender, string message, bool isStreaming = false, Action<Action<string>> setUpdateContent = null)
        {
            if (string.IsNullOrWhiteSpace(message) && !isStreaming) return null;
            Control contentControl;

            if (sender == "AI")
            {
                // 初始化一个新的webview2，设置基本属性
                // webview2是一个浏览器控件，用于在winform中渲染html，js，markdown等内容
                var webView = new WebView2
                {
                    Width = conversationPanel.Width - 150,
                    Margin = new Padding(0),
                    Height = 50,
                    DefaultBackgroundColor = Color.White
                };
                // 异步初始化浏览器内核
                await webView.EnsureCoreWebView2Async();
                if (webView.CoreWebView2 != null)
                {
                    // 添加节流变量
                    DateTime lastScrollTime = DateTime.MinValue;
                    int lastHeight = 0;
                    // (s, a) => {}是一个lambda表达式，其中(s, a)是参数
                    // s是触发事件的对象，a是事件参数
                    // webmessagereceived是一个监听器，监听js向c#发送消息的回调
                    // +=表示触发webmessagereceived之后，这个匿名函数就会执行
                    // js发送消息的逻辑写在了下面的doc中
                    webView.CoreWebView2.WebMessageReceived += (s, a) =>
                    {
                        // int.tryparse()将js向c#发送的内容解析为一个整数，如果成功就传给height并返回true
                        if (int.TryParse(a.WebMessageAsJson, out int height))
                        {
                            // 只有高度变化显著时才更新
                            if (Math.Abs(height - lastHeight) > 10)
                            {
                                lastHeight = height;
                                // 定义一个叫update的匿名函数，类型是action，表示没有参数也没有返回值的一个函数
                                // 之后运行update()即可执行这个函数
                                Action update = () =>
                                {
                                    webView.Height = Math.Max(height, 50);

                                    // 节流滚动：限制滚动频率
                                    var now = DateTime.Now;
                                    if ((now - lastScrollTime).TotalMilliseconds > 150) // 150ms节流
                                    {
                                        lastScrollTime = now;

                                        // 延迟滚动，让界面有时间更新
                                        Task.Delay(30).ContinueWith(_ =>
                                        {
                                            try
                                            {
                                                if (webView.IsHandleCreated && !webView.IsDisposed)
                                                {
                                                    webView.Invoke(new Action(() =>
                                                    {
                                                        try
                                                        {
                                                            SmartAutoScroll(webView);
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            Debug.WriteLine("SmartAutoScroll failed: " + ex.Message);
                                                        }
                                                    }));
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                Debug.WriteLine("Async scroll error: " + ex.Message);
                                            }
                                        });
                                    }
                                };

                                if (webView.InvokeRequired)
                                    webView.Invoke(update);
                                else
                                    update();
                            }
                        }
                    };

                    // 优化HTML结构，减少重排
                    var pipeline = new MarkdownPipelineBuilder()
                                        .UseAdvancedExtensions()  // ✅ 启用 GFM 表格等支持
                                        .Build();

                    string html = Markdown.ToHtml(message ?? "", pipeline);
                    string doc = $@"
<html>
<head>
    <script>
        window.MathJax = {{
            tex: {{
                inlineMath: [['$','$'], ['\\(','\\)']],
                displayMath: [['$$','$$'], ['\\[','\\]']]
            }},
            svg: {{
                fontCache: 'global'
            }}
        }};
    </script>
    <script src='https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js'></script>
    <style>
        html, body {{
            margin: 0;
            padding: 0;
            font-family: 微软雅黑;
            font-size: {htmlsize};
            background-color: #FFFFFF;
            line-height: 1.6;
            word-wrap: break-word;
            /* ✅ 关键：禁用滚动条 */
            overflow: hidden !important;
            overflow-x: hidden !important;
            overflow-y: hidden !important;
        }}
        .content {{
            padding: 8px;
            min-height: 40px; /* ✅ 设置最小高度 */
            /* ✅ 确保内容不会溢出 */
            overflow: hidden;
            word-break: break-word;
        }}
        /* ✅ 新增：表格样式 - 关键修复 */
        table {{
            border-collapse: collapse;
            width: 100%;
            max-width: 100%;
            margin: 10px 0;
            background-color: #fff;
            /* ✅ 防止表格溢出 */
            table-layout: fixed;
            word-wrap: break-word;
        }}
        
        th, td {{
            border: 1px solid #ddd;
            padding: 8px 12px;
            text-align: left;
            vertical-align: top;
            /* ✅ 确保单元格内容不溢出 */
            overflow: hidden;
            word-break: break-word;
            hyphens: auto;
        }}
        
        th {{
            background-color: #f2f2f2;
            font-weight: bold;
        }}
        
        /* ✅ 响应式表格 */
        @media screen and (max-width: 600px) {{
            table {{
                font-size: 12px;
            }}
            th, td {{
                padding: 6px 8px;
            }}
        }}
        
        /* ✅ 处理长内容 */
        .content table {{
            overflow-wrap: break-word;
            word-wrap: break-word;
        }}
        /* ✅ 隐藏所有可能的滚动条 */
        ::-webkit-scrollbar {{
            display: none !important;
            width: 0 !important;
            height: 0 !important;
        }}
        * {{
            scrollbar-width: none !important;
            -ms-overflow-style: none !important;
        }}
    </style>
</head>
<body>
    <div id='content' class='content'>{html}</div>
    <script>
        let updateTimer = null;
        let isUpdating = false;
        let lastReportedHeight = 0;
        
        function debounceHeightUpdate() {{
            if (updateTimer) {{
                clearTimeout(updateTimer);
            }}
            updateTimer = setTimeout(() => {{
                const content = document.getElementById('content');
                
                // ✅ 更准确的高度计算
                const height = Math.max(
                    content.scrollHeight + 16, // padding
                    content.offsetHeight + 16,
                    document.documentElement.scrollHeight,
                    document.body.scrollHeight,
                    50 // 最小高度
                );
                
                // ✅ 避免频繁发送相同高度
                if (Math.abs(height - lastReportedHeight) > 3) {{
                    lastReportedHeight = height;
                    window.chrome.webview.postMessage(height);
                }}
            }}, 30); // 减少防抖时间，更快响应
        }}
        
        function setContent(html) {{
            if (isUpdating) return;
            isUpdating = true;
            
            const content = document.getElementById('content');
            
            // ✅ 使用更高效的DOM更新方式
            requestAnimationFrame(() => {{
                try {{
                    content.innerHTML = html;
                    
                    // ✅ 立即计算并报告高度，避免滚动条闪现
                    const immediateHeight = Math.max(
                        content.scrollHeight + 16,
                        content.offsetHeight + 16,
                        50
                    );
                    
                    if (Math.abs(immediateHeight - lastReportedHeight) > 3) {{
                        lastReportedHeight = immediateHeight;
                        window.chrome.webview.postMessage(immediateHeight);
                    }}
                    
                    if (window.MathJax) {{
                        MathJax.typesetPromise([content]).then(() => {{
                            debounceHeightUpdate();
                            isUpdating = false;
                        }}).catch(() => {{
                            debounceHeightUpdate();
                            isUpdating = false;
                        }});
                    }} else {{
                        debounceHeightUpdate();
                        isUpdating = false;
                    }}
                }} catch (e) {{
                    console.error('Content update error:', e);
                    isUpdating = false;
                }}
            }});
        }}
        
        // 初始化
        document.addEventListener('DOMContentLoaded', function() {{
            if (window.MathJax) {{
                MathJax.typesetPromise().then(() => {{
                    debounceHeightUpdate();
                }});
            }} else {{
                debounceHeightUpdate();
            }}
        }});
        
        // ✅ 优化的观察器，减少触发频率
        const observer = new MutationObserver((mutations) => {{
            let shouldUpdate = false;
            mutations.forEach(mutation => {{
                if (mutation.type === 'childList' || 
                    (mutation.type === 'characterData' && mutation.target.textContent.length > 0)) {{
                    shouldUpdate = true;
                }}
            }});
            
            if (shouldUpdate) {{
                debounceHeightUpdate();
            }}
        }});
        
        observer.observe(document.getElementById('content'), {{
            childList: true,
            subtree: true,
            characterData: true
        }});
        
        // ✅ 页面加载完成后立即报告高度
        window.addEventListener('load', function() {{
            debounceHeightUpdate();
        }});
    </script>
</body>
</html>";
                    

                    webView.CoreWebView2.NavigateToString(doc);

                    if (setUpdateContent != null)
                    {
                        // 添加更新节流
                        DateTime lastUpdateTime = DateTime.MinValue;
                        string lastContent = "";

                        Action<string> updateContent = (newContent) =>
                        {
                            // 内容去重
                            if (newContent == lastContent) return;
                            lastContent = newContent;

                            // 更新频率控制
                            var now = DateTime.Now;
                            if ((now - lastUpdateTime).TotalMilliseconds < 100) // 100ms内不重复更新
                            {
                                return;
                            }
                            lastUpdateTime = now;

                            try
                            {
                                string htmlUpdate = Markdig.Markdown.ToHtml(newContent)
                                    .Replace("`", "\\`")
                                    .Replace("\\", "\\\\")
                                    .Replace("'", "\\'")
                                    .Replace("\r\n", "\\n")
                                    .Replace("\n", "\\n");

                                if (webView.IsHandleCreated && !webView.IsDisposed)
                                {
                                    if (webView.InvokeRequired)
                                    {
                                        webView.Invoke(new Action(() =>
                                        {
                                            webView.CoreWebView2?.ExecuteScriptAsync($"setContent(`{htmlUpdate}`);");
                                        }));
                                    }
                                    else
                                    {
                                        webView.CoreWebView2?.ExecuteScriptAsync($"setContent(`{htmlUpdate}`);");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                // 记录错误但不中断流程
                                System.Diagnostics.Debug.WriteLine($"WebView update error: {ex.Message}");
                            }
                        };

                        setUpdateContent(updateContent);
                    }
                }
                conversationPanel.Controls.Add(webView);
                contentControl = webView;
            }
            else
            {
                contentControl = new TextBox
                {
                    Text = message,
                    Font = new System.Drawing.Font("微软雅黑", fontsize),
                    MaximumSize = new Size(conversationPanel.Width - 150, 0),
                    Padding = new Padding(10),
                    BackColor = Color.LightBlue,
                    Margin = new Padding(5),
                    BorderStyle = BorderStyle.None, // 看起来像Label
                    ReadOnly = true,                // 只读
                    Multiline = true,              // 支持多行
                    ScrollBars = ScrollBars.None,  // 不显示滚动条
                    TabStop = false,               // 不参与Tab导航
                    Cursor = Cursors.IBeam         // 文本光标
                };
                Size textSize = TextRenderer.MeasureText(message, contentControl.Font, contentControl.MaximumSize, TextFormatFlags.WordBreak);
                contentControl.Height = textSize.Height + 10; // +10 是额外填充防止被裁切
                // 设置实际宽度和高度
contentControl.Width = Math.Min(textSize.Width + contentControl.Padding.Horizontal, conversationPanel.Width - 150);
contentControl.Height = textSize.Height + contentControl.Padding.Vertical;
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
        private void ConversationPanel_Resize(object sender, EventArgs e)
        {
            // 面板大小改变时，如果在底部则保持在底部
            if (isAutoScrollEnabled && IsNearBottom())
            {
                Task.Delay(100).ContinueWith(_ =>
                {
                    this.Invoke(new Action(() => SmartAutoScroll()));
                });
            }
        }
        // 非流式输出
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
<body style='font-family:微软雅黑;font-size:{htmlsize};background-color:#FFFFFF;margin:0;padding:0;'>
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
                    Font = new System.Drawing.Font("微软雅黑", fontsize),
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
        #endregion

        // 流式调用
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
                Content = "你的名字是\"电气设计助手\"，是一个电气设计领域的AutoCAD助手，请回答有关电气设计领域的问题并返回原始的Markdown格式"
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
        // 非流式调用
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
  
    }
}
