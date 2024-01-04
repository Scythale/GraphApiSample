using Enovatio.Emails.Providers;

namespace GraphApiSample
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            using (var client = new ExchangeOnlineMailClient())
            {
                await foreach (var message in client.GetMessagesFromInboxAsync())
                {
                    Console.WriteLine(message.Subject);
                }
            }
        }
    }
}
