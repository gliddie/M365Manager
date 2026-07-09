namespace M365Manager.Core.PowerShell;

/// <summary>Raised when a hosted PowerShell command writes to the error stream.</summary>
public sealed class PowerShellException : Exception
{
    public PowerShellException(string message) : base(message) { }
}
