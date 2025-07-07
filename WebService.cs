using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CoDesignStudy.Cad.PlugIn
{
    public class WebServicePanel : UserControl
    {
        private WebView2 webView;

        public WebServicePanel(string url)
        {
            InitializeComponent(url);
        }

        private void InitializeComponent(string url)
        {
            this.Dock = DockStyle.Fill;
            webView = new WebView2
            {
                Dock = DockStyle.Fill
            };
            this.Controls.Add(webView);

            // 初始化并加载网页
            webView.EnsureCoreWebView2Async().ContinueWith(_ =>
            {
                if (webView.CoreWebView2 != null)
                {
                    webView.CoreWebView2.Navigate(url);
                }
            }, TaskScheduler.FromCurrentSynchronizationContext());
        }
    }
}