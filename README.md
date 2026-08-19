# Plugboard Core Plugins

Official connector and service plugins for [Plugboard](https://github.com/PrismShell/Plugboard).

## Included Plugins

| Plugin | Type | Description |
|--------|------|-------------|
| **Bbg** | Connector | BBG Terminal data — BDP, BDS, BDH, field/security search via bbcomm |
| **Cmp** | Connector | BBG CMP (CMBS/structured analytics) |
| **Excel** | Connector | Read/write Excel workbooks via COM automation |
| **Outlook** | Connector | Email, folders, calendar, contacts via Outlook COM |
| **Pdf** | Connector | PDF generation and table extraction |
| **Files** | Connector | Bounded filesystem access |
| **Sql** | Connector | SQL Server database queries |
| **Sqlite** | Connector | Local SQLite record store with changelog + newest-wins file-share sync |
| **Xlsx** | Connector | Read .xlsx files off disk/shares - no Excel, no COM |
| **Ping** | Connector | Simple demo/health check connector |
| **MeshCheck** | Service | Self-test for the plugin mesh |
| **Store** | Service | In-memory shared state store (namespaced key/JSON CRUD + subset select) |
| **MarketCache** | Service | In-memory TTL cache for vendor market data |

## Building

Requires .NET 8 SDK.

`ash
# Build all plugins
dotnet build

# Or use the Plugboard build script
cd path/to/plugboard
.\tools\build-plugins.ps1
`

## Deployment

Built plugins are deployed to Plugboard's `plugins` directory. See the [Plugboard README](https://github.com/PrismShell/Plugboard) for plugin loading and signing.

## License

AGPL-3.0
