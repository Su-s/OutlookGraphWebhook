using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace OutlookGraphWebhook.Controllers;

[ApiController]
[Route("api/subscribe")]
public class SubscriptionController : ControllerBase
{
    private readonly ILogger<SubscriptionController> _logger;
    private readonly GraphServiceClient _graphServiceClient;
    private readonly IConfiguration _configuration;

    public SubscriptionController(ILogger<SubscriptionController> logger, GraphServiceClient graphServiceClient,
        IConfiguration configuration)
    {
        _logger = logger;
        _graphServiceClient = graphServiceClient;
        _configuration = configuration;
    }

    [HttpPost]
    public async Task<IResult> Post(SubscriptionRequest request)
    {
        try
        {
            var filterQuery = "";
            if (!string.IsNullOrEmpty(request.FilterQuery))
                filterQuery = "&$filter=" + request.FilterQuery;
            var subscription = new Subscription
            {
                ChangeType = "created",
                NotificationUrl = $"{_configuration.GetValue<string>("Ngrok")}/api/listen",
                LifecycleNotificationUrl = $"{_configuration.GetValue<string>("Ngrok")}/api/lifecycle",
                Resource =
                    $"/me/mailfolders('{request.MailBox}')/messages?$select=Subject,bodyPreview,importance,receivedDateTime,from{filterQuery}",
                ExpirationDateTime = DateTimeOffset.UtcNow.AddMinutes(4230),
                ClientState = Guid.NewGuid().ToString(),
                LatestSupportedTlsVersion = "v1_2"
            };

            var createdSubscription = await _graphServiceClient.Subscriptions
                .PostAsync(subscription);

            return Results.Ok($"Subscription created with ID: {createdSubscription.Id}");
        }
        catch (Exception ex)
        {
            return Results.BadRequest($"Error creating subscription: {ex.Message}");
        }
    }

    [HttpDelete]
    public async Task<IResult> Delete(string subscriptionId)
    {
        try
        {
            await _graphServiceClient.Subscriptions[subscriptionId]
                .DeleteAsync();

            return Results.Ok($"Subscription {subscriptionId} deleted.");
        }
        catch (Exception ex)
        {
            return Results.BadRequest($"Error deleting subscription: {ex.Message}");
        }
    }
}

public class SubscriptionRequest
{
    public string MailBox { get; set; }
    public string FilterQuery { get; set; }
}