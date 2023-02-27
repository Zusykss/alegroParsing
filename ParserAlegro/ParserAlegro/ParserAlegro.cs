using PuppeteerSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParserAlegro
{
    public class ParserAlegro
    {

        private readonly HttpClient client = new HttpClient();
        public Browser _browser;
        public Page _page;
        protected string Proxy = "", Login = "", Password = "";
        
        LaunchOptions options;

        public void SerProxy(string Proxy)
        {
            if (Proxy.Length == 0) return;
            var path = File.ReadAllText("path.txt");
            var list = Proxy.Split('@');
            this.Proxy = list[0];
            var userPass = list[1].Split(':');
            Login = userPass[0];
            Password = userPass[1];
            options = new LaunchOptions
            {
                Headless = false,
                Args = new[]{
            "--start-maximized",
            $"--proxy-server=http://{this.Proxy}"},
            IgnoredDefaultArgs = new string[]
            {
                "--enable-automation"
            },
               // ExecutablePath = @"chrome-win\chrome.exe"
               ExecutablePath = path
            };
        }
        public async Task InitBrowserAsync()
        {
            _browser = await Puppeteer.LaunchAsync(options);
        }

        public async Task<Page> GetPageAsync()
        {
            return await _browser.NewPageAsync();
        }
        public  async Task CloseBrowser()
        {
            await _browser.CloseAsync();
        }

        public async Task<string> BrowserLoader(string url, int count)
        {
            string text = "";
            try
            {
                await InitBrowserAsync();
                _page = await GetPageAsync();
                await _page.AuthenticateAsync(new PuppeteerSharp.Credentials() { Username = Login, Password = Password });

                var responce = await _page.GoToAsync(url);
              
                text = await responce.TextAsync();
            }
            catch (Exception ex)
            {
               text = "ERROR BrowserLoader " + ex.Message;
                //var responce = await _page.GoToAsync("https://github.com/");

                File.AppendAllText("logs.txt", DateTime.Now.ToString("MM/dd/yyyy HH:mm") + " - " + text + $" - {url}\r\n", Encoding.UTF8);
                if (ex.Message.Contains("Timeout of") && count < 5)
                {
                    await CloseBrowser();
                    count++;
                    return await BrowserLoader(url, count);
                }
            }
            await CloseBrowser();

            return text;
        }
    }
}
