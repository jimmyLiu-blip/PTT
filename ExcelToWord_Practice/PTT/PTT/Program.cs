using System;
using System.Net.Http;
using System.Threading.Tasks;
using HtmlAgilityPack;

class Program
{
    static async Task Main()
    {
        string url = "https://www.ptt.cc/bbs/hotboards.html";

        using (HttpClient client = new HttpClient())
        {
            // PTT 需要年滿18歲的 Cookie 才能訪問部分看板
            client.DefaultRequestHeaders.Add("Cookie", "over18=1");

            Console.WriteLine($"正在抓取：{url}\n");

            string html = await client.GetStringAsync(url);
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);

            // 選取所有熱門看板標題
            var boardNodes = doc.DocumentNode.SelectNodes("//div[@class='board-name']");

            if (boardNodes != null)
            {
                int index = 1;
                foreach (var node in boardNodes)
                {
                    Console.WriteLine($"{index++}. {node.InnerText}");
                }
            }
            else
            {
                Console.WriteLine("⚠️ 沒有找到任何看板標題，請確認XPath或網頁結構是否變更。");
            }
        }

        Console.WriteLine("\n✅ 抓取完成！");
    }
}
