using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.Messages.Item.Move;

namespace Enovatio.Emails.Providers;

internal class ExchangeOnlineMailClient : IDisposable
{
    private string _userName = "";
    private string _tenantId = "841fea1b-c112-4e30-a6e4-ab424d9b4af2"; // enovatio
    private string _clientId = "a1a937ea-da97-4a71-b380-eeaa18639bef";
    private string _clientSecret = "";

    private GraphServiceClient _graphServiceClient;
    private MailFolder _invalidFolder;
    private MailFolder _succeededFolder;

    public async IAsyncEnumerable<Message> GetMessagesFromInboxAsync()
    {
        await InitializeClient().ConfigureAwait(false);
        var messages = await _graphServiceClient.Users[_userName].MailFolders["Inbox"].Messages.GetAsync();
        foreach (var message in messages?.Value)
        {
            var attachments = await _graphServiceClient.Users[_userName].Messages[message.Id].Attachments.GetAsync();
            message.Attachments = attachments?.Value;
            yield return message;
        }
    }

    public async Task MoveToInvalid(string uid)
    {
        await InitializeClient().ConfigureAwait(false);
        await _graphServiceClient.Users[_userName].Messages[uid].Move.PostAsync(new MovePostRequestBody
        {
            DestinationId = _invalidFolder.Id
        });
    }

    public async Task MoveToSucceeded(string uid)
    {
        await InitializeClient().ConfigureAwait(false);
        await _graphServiceClient.Users[_userName].Messages[uid].Move.PostAsync(new MovePostRequestBody
        {
            DestinationId = _succeededFolder.Id
        });
    }

    private async Task<MailFolder> CreateOrGetFolderAsync(string folderName)
    {
        var folders = await _graphServiceClient.Users[_userName].MailFolders.GetAsync(p =>
        {
            p.QueryParameters.Select = new[] { "id", "displayName" };
            p.QueryParameters.Filter = $"displayName eq '{folderName}'";
        });

        if (folders?.Value?.Count == 1)
        {
            return folders.Value.Single();
        }

        return await _graphServiceClient.Users[_userName].MailFolders.PostAsync(new MailFolder
        {
            DisplayName = folderName,
            IsHidden = false
        });

    }


    private async Task InitializeClient()
    {
        if (_graphServiceClient == null)
        {
            _graphServiceClient = new GraphServiceClient(new ClientSecretCredential(_tenantId, _clientId, _clientSecret));

            _invalidFolder = await CreateOrGetFolderAsync("Invalid");
            _succeededFolder = await CreateOrGetFolderAsync("Successful");
        }
    }

    public void Dispose()
    {
        _graphServiceClient.Dispose();
    }
}