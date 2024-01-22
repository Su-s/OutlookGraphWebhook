using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace OutlookGraphWebhook.Controllers;

[ApiController]
[Route("api/listen")]
public class ListenController : ControllerBase
{
    private readonly ILogger<ListenController> _logger;
    private readonly GraphServiceClient _graphServiceClient;
    private readonly IConfiguration _configuration;

    public ListenController(ILogger<ListenController> logger, GraphServiceClient graphServiceClient,
        IConfiguration configuration)
    {
        _logger = logger;
        _graphServiceClient = graphServiceClient;
        _configuration = configuration;
    }

    [HttpPost]
    public async Task<IResult> Post([FromBody] string content)
    {
        // Parse the received notifications.
        var notifications = JsonConvert.DeserializeObject<ChangeNotificationCollectionResponse>(content);

        foreach (var notification in notifications.Value)
        {
            // Handle the change notification (this will depend on your requirements).
            // For example, you might fetch the email message and log it, send a notification, etc.
            var message = await _graphServiceClient.Me.Messages[notification.SubscriptionId.ToString()].GetAsync();
            Console.WriteLine($"Received message: {message.Subject}");
        }

        // Return a 202 status code to acknowledge receipt of the notification.
        // Response.StatusCode = 202;
        return Results.Ok($"Subscription listened.");
    }
}