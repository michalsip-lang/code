using Microsoft.EntityFrameworkCore;
using System.Data;

namespace AkcniPlan.Data;

public static class SchemaUpdater
{
    public static void Apply(AppDbContext db)
    {
        EnsureColumnArea(db);
        EnsureColumnSourceMessageId(db);
    }

    private static void EnsureColumnArea(AppDbContext db)
    {
        if (ColumnExists(db, "Tasks", "Area"))
        {
            return;
        }

        db.Database.ExecuteSqlRaw("ALTER TABLE Tasks ADD COLUMN Area INTEGER NOT NULL DEFAULT 4;");
    }

    private static void EnsureColumnSourceMessageId(AppDbContext db)
    {
        if (ColumnExists(db, "Tasks", "SourceMessageId"))
        {
            return;
        }

        db.Database.ExecuteSqlRaw("ALTER TABLE Tasks ADD COLUMN SourceMessageId TEXT NULL;");
    }

    private static bool ColumnExists(AppDbContext db, string tableName, string columnName)
    {
        using var command = db.Database.GetDbConnection().CreateCommand();
        command.CommandText = $"PRAGMA table_info({tableName});";

        if (command.Connection is not null && command.Connection.State != ConnectionState.Open)
        {
            command.Connection.Open();
        }

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            var current = reader[1]?.ToString();
            if (string.Equals(current, columnName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }
}