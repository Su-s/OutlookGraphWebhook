using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;

namespace OutlookGraphWebhook.Controllers;

[ApiController]
[Route("api/lifecycle")]
public class LifecycleController : ControllerBase
{
    private readonly ILogger<LifecycleController> _logger;
    private readonly GraphServiceClient _graphServiceClient;

    public LifecycleController(ILogger<LifecycleController> logger, GraphServiceClient graphServiceClient)
    {
        _logger = logger;
        _graphServiceClient = graphServiceClient;
    }

    [HttpPost]
    public async Task<IResult> Post([FromBody] string subscriptionId)
    {
        try
        {
            var subscription = await _graphServiceClient.Subscriptions[subscriptionId]
                .GetAsync();

            subscription.ExpirationDateTime = DateTimeOffset.UtcNow.AddMinutes(4230);

            await _graphServiceClient.Subscriptions[subscription.Id]
                .PatchAsync(subscription);

            return Results.Ok($"Subscription {subscription.Id} renewed.");
        }
        catch (Exception ex)
        {
            return Results.BadRequest($"Error renewing subscription: {ex.Message}");
        }
    }
}