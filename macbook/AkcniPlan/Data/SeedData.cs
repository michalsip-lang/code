using AkcniPlan.Models;

namespace AkcniPlan.Data;

public static class SeedData
{
    public static void Initialize(AppDbContext db)
    {
        if (db.Tasks.Any())
        {
            return;
        }

        var now = DateTime.UtcNow;
        var tasks = new List<TaskItem>
        {
            new()
            {
                Title = "Pripravit tydenni plan",
                Description = "Rozdelit praci na dny a odhadnout kapacitu.",
                DueDate = now.Date.AddDays(1),
                EstimatedHours = 2,
                ActualHours = 0,
                Importance = 4,
                Area = TaskArea.Sdp,
                Status = TaskItemStatus.Todo,
                CreatedAt = now.AddDays(-2)
            },
            new()
            {
                Title = "Dokoncit report vykonnosti",
                Description = "Doplnit KPI a komentare do mesicniho reportu.",
                DueDate = now.Date,
                EstimatedHours = 4,
                ActualHours = 1.5,
                Importance = 5,
                Area = TaskArea.Svp,
                Status = TaskItemStatus.InProgress,
                CreatedAt = now.AddDays(-4)
            },
            new()
            {
                Title = "Vyresit opozdene faktury",
                Description = "Kontaktovat financni oddeleni a potvrdit plan uhrad.",
                DueDate = now.Date.AddDays(-1),
                EstimatedHours = 3,
                ActualHours = 0,
                Importance = 5,
                Area = TaskArea.Jine,
                Status = TaskItemStatus.Todo,
                CreatedAt = now.AddDays(-3)
            },
            new()
            {
                Title = "Uzavrit dokumentaci projektu",
                Description = "Finalni revize a archivace dokumentu.",
                DueDate = now.Date.AddDays(-3),
                EstimatedHours = 5,
                ActualHours = 6,
                Importance = 3,
                Area = TaskArea.Bozp,
                Status = TaskItemStatus.Done,
                CreatedAt = now.AddDays(-12),
                CompletedAt = now.AddDays(-2)
            }
        };

        db.Tasks.AddRange(tasks);

        var tags = new List<TaskTag>
        {
            new() { Name = "Finance" },
            new() { Name = "Planovani" },
            new() { Name = "Reporting" }
        };
        db.Tags.AddRange(tags);
        db.SaveChanges();

        db.TaskItemTags.AddRange(
            new TaskItemTag { TaskItemId = tasks[0].Id, TaskTagId = tags[1].Id },
            new TaskItemTag { TaskItemId = tasks[1].Id, TaskTagId = tags[2].Id },
            new TaskItemTag { TaskItemId = tasks[2].Id, TaskTagId = tags[0].Id }
        );

        db.TaskDependencies.Add(
            new TaskDependency { TaskItemId = tasks[1].Id, DependsOnTaskItemId = tasks[0].Id }
        );

        db.SaveChanges();
    }
}
