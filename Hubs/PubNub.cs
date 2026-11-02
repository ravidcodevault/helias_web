using PubnubApi;

namespace OYA_Web_Application.Hubs;

public class PubNub
{
    private readonly IConfiguration _configuration;
    public const string ChannelName = "OYA_SYNC_UPDATE";
    private const string PublisherUserId = "OYA_WEB_USER";
    public string PublishKey => _configuration["PubNub:PublishKey"]
        ?? throw new InvalidOperationException("Configure PubNub keys in private runtime settings.");
    public string SubscribeKey => _configuration["PubNub:SubscribeKey"]
        ?? throw new InvalidOperationException("Configure PubNub keys in private runtime settings.");

    public PubNub(IConfiguration configuration) => _configuration = configuration;

    public async Task<string> CreateReadTokenAsync(string userId, CancellationToken cancellationToken = default)
        => await GrantTokenAsync(userId, read: true, write: false, cancellationToken);

    private async Task<string> GrantTokenAsync(string userId, bool read, bool write,
        CancellationToken cancellationToken)
    {
        var secretKey = _configuration["PubNub:SecretKey"];
        if (string.IsNullOrWhiteSpace(secretKey))
            throw new InvalidOperationException("PubNub authorization is not configured. Set PubNub:SecretKey in server-side secrets for this keyset.");

        var issuer = new Pubnub(new PNConfiguration(new UserId(PublisherUserId))
        {
            PublishKey = PublishKey,
            SubscribeKey = SubscribeKey,
            SecretKey = secretKey
        });
        try
        {
            var grant = await issuer.GrantToken().TTL(15).AuthorizedUuid(userId)
                .Resources(new PNTokenResources
                {
                    Channels = new Dictionary<string, PNTokenAuthValues>
                    {
                        [ChannelName] = new PNTokenAuthValues { Read = read, Write = write }
                    }
                }).ExecuteAsync().WaitAsync(TimeSpan.FromSeconds(30), cancellationToken);
            if (grant.Status.Error || string.IsNullOrWhiteSpace(grant.Result?.Token))
                throw new InvalidOperationException("PubNub could not authorize sync access. Verify the server-side SecretKey belongs to this keyset.");
            return grant.Result.Token;
        }
        finally { issuer.Destroy(); }
    }

    public async Task SendMessageAsync(Dictionary<string, string> message,
        CancellationToken cancellationToken = default)
    {
        var token = await GrantTokenAsync(PublisherUserId, read: false, write: true, cancellationToken);
        var publisher = new Pubnub(new PNConfiguration(new UserId(PublisherUserId))
        {
            PublishKey = PublishKey,
            SubscribeKey = SubscribeKey
        });
        try
        {
            publisher.SetAuthToken(token);
            var result = await publisher.Publish().Channel(ChannelName).Message(message)
                .ExecuteAsync().WaitAsync(TimeSpan.FromSeconds(30), cancellationToken);
            if (result.Status.Error)
                throw new InvalidOperationException("PubNub rejected the sync command. Check the keyset and channel permissions.");
        }
        finally { publisher.Destroy(); }
    }
}
