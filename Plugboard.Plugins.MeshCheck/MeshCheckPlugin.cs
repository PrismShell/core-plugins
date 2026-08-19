using Plugboard.Contracts;

namespace Plugboard.Plugins.MeshCheck;

// A runnable SERVICE with no external dependencies: it composes the ping
// CONNECTOR to prove the mesh works end to end on any machine (no Terminal
// needed). Calling svc/meshcheck makes the host dispatch to ping/hello
// in-process via req.Call and wraps the result — one unit calling another.
public sealed class MeshCheckPlugin : IPlugin
{
    public string Name => "meshcheck";

    public void Register(IEndpointRegistry r) =>
        r.Map("GET", "svc/meshcheck", async req =>
        {
            var pong = await req.Call("ping/hello", null);   // connector call, by name
            return (object?)new
            {
                service  = "meshcheck",
                composed = pong,          // whatever ping/hello returned
                depth    = req.Depth
            };
        },
        new RouteInfo("Mesh self-test",
            "Calls the ping connector via req.Call to prove connector→service composition."));
}
