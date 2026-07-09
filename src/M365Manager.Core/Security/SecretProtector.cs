using System.Security.Cryptography;
using System.Text;

namespace M365Manager.Core.Security;

/// <summary>
/// Encrypts/decrypts secrets (SQL password, M365 client secret) using Windows DPAPI
/// scoped to the current user. Ciphertext is only decryptable by the same Windows
/// user on the same machine, and is stored base64-encoded in the settings file.
/// </summary>
public static class SecretProtector
{
    public static string? Protect(string? plaintext)
    {
        if (string.IsNullOrEmpty(plaintext))
            return plaintext;

        var bytes = Encoding.UTF8.GetBytes(plaintext);
        var encrypted = ProtectedData.Protect(bytes, optionalEntropy: null, DataProtectionScope.CurrentUser);
        return Convert.ToBase64String(encrypted);
    }

    public static string? Unprotect(string? ciphertext)
    {
        if (string.IsNullOrEmpty(ciphertext))
            return ciphertext;

        try
        {
            var encrypted = Convert.FromBase64String(ciphertext);
            var bytes = ProtectedData.Unprotect(encrypted, optionalEntropy: null, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(bytes);
        }
        catch
        {
            // Value was not DPAPI-encrypted (e.g. legacy/manual edit) or cannot be decrypted.
            return ciphertext;
        }
    }
}
