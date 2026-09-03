using Microsoft.Data.SqlClient;

namespace M365Manager.ViewModels;

/// <summary>
/// Turns an exception into something an admin can act on.
///
/// EF Core wraps every database failure in a DbUpdateException whose own message is the useless
/// "An error occurred while saving the entity changes. See the inner exception for details." -
/// the actual cause (missing table, duplicate key, permission denied) only lives in the inner
/// exception, so showing ex.Message alone tells the operator nothing.
/// </summary>
public static class ErrorText
{
    public static string Describe(Exception ex)
    {
        var root = Unwrap(ex);

        if (root is SqlException sql)
        {
            var hint = sql.Number switch
            {
                // "Invalid object name" - the table hasn't been created yet.
                208 => " The table is missing - run the matching script from src/M365Manager.Data/Scripts against the database (06_RoomResources.sql for room sites).",
                // Unique/primary key violation.
                2601 or 2627 => " An entry with that key already exists.",
                // Login/permission problems.
                229 or 300 or 262 => " The SQL login is missing permissions on this object.",
                _ => "",
            };
            return sql.Message + hint;
        }

        return root.Message;
    }

    /// <summary>Walks to the innermost exception - that's the one carrying the real cause.</summary>
    private static Exception Unwrap(Exception ex)
    {
        var current = ex;
        while (current.InnerException is not null)
            current = current.InnerException;
        return current;
    }
}
