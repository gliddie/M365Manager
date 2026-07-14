namespace M365Manager.Core.M365;

public enum ConnectorState { Connecting, Connected, Failed }

/// <summary>One connector's state during a <see cref="IStartupConnectionService.ConnectAllAsync"/> run.</summary>
public sealed record ConnectorProgress(string DisplayName, ConnectorState State, string? Error = null);

public interface IStartupConnectionService
{
    /// <summary>
    /// Connects every registered <see cref="IM365Connector"/> (Graph, Exchange Online, and any
    /// future connector), in <see cref="IM365Connector.Order"/> order, skipping ones that are
    /// already connected. One connector failing doesn't stop the rest - e.g. Exchange being
    /// unreachable shouldn't prevent a future Teams/SharePoint connector from still signing in.
    /// </summary>
    Task ConnectAllAsync(Action<ConnectorProgress>? onProgress = null, Action<string>? onPrompt = null, CancellationToken ct = default);
}

public sealed class StartupConnectionService : IStartupConnectionService
{
    private readonly IEnumerable<IM365Connector> _connectors;

    public StartupConnectionService(IEnumerable<IM365Connector> connectors)
    {
        _connectors = connectors;
    }

    public async Task ConnectAllAsync(Action<ConnectorProgress>? onProgress = null, Action<string>? onPrompt = null, CancellationToken ct = default)
    {
        foreach (var connector in _connectors.OrderBy(c => c.Order))
        {
            if (connector.IsConnected)
            {
                onProgress?.Invoke(new ConnectorProgress(connector.DisplayName, ConnectorState.Connected));
                continue;
            }

            onProgress?.Invoke(new ConnectorProgress(connector.DisplayName, ConnectorState.Connecting));
            try
            {
                await connector.ConnectAsync(onPrompt, ct);
                onProgress?.Invoke(new ConnectorProgress(connector.DisplayName, ConnectorState.Connected));
            }
            catch (Exception ex)
            {
                onProgress?.Invoke(new ConnectorProgress(connector.DisplayName, ConnectorState.Failed, ex.Message));
            }
        }
    }
}
