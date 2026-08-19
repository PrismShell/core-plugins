using Plugboard.Contracts;

namespace Plugboard.Plugins.Ping;

// Trivial proof plugin: validates the discover -> verify -> load -> register ->
// serve loop end to end without any external dependencies. Real capability
// plugins (Bbg, Excel, Outlook) follow the same shape.
public sealed class PingPlugin : IPlugin
{
    public string Name => "ping";

    public void Register(IEndpointRegistry registry)
    {
        registry.Map("GET", "ping/hello", _ =>
            Task.FromResult<object?>(new { message = "pong", at = DateTime.UtcNow.ToString("o") }));

        registry.Map("POST", "ping/echo", req =>
            Task.FromResult<object?>(new { youSent = req.Body, query = req.Query }));
    }
}
