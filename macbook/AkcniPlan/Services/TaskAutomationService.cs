using AkcniPlan.Data;
using AkcniPlan.Models;
using Microsoft.EntityFrameworkCore;

namespace AkcniPlan.Services;

public class TaskAutomationService
{
    private readonly AppDbContext _db;
    private readonly PriorityScoringService _scoring;

    public TaskAutomationService(AppDbContext db, PriorityScoringService scoring)
    {
        _db = db;
        _scoring = scoring;
    }

    public async Task<int> GenerateFromTextAsync(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return 0;
        }

        var lines = input
            .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Distinct()
            .ToList();

        if (lines.Count == 0)
        {
            return 0;
        }

        var existing = await _db.Tasks
            .Select(t => t.Title.ToLower())
            .ToListAsync();

        var created = 0;
        foreach (var line in lines)
        {
            if (existing.Contains(line.ToLower()))
            {
                continue;
            }

            var task = new TaskItem
            {
                Title = line,
                Description = "Automaticky vygenerovaný úkol z textového vstupu.",
                DueDate = DateTime.UtcNow.Date.AddDays(2),
                EstimatedHours = 1,
                Importance = 3,
                Area = InferArea(line, string.Empty),
                Status = TaskItemStatus.Todo,
                CreatedAt = DateTime.UtcNow
            };
            task.PriorityScore = _scoring.CalculateScore(task, DateTime.UtcNow);
            _db.Tasks.Add(task);
            created++;
        }

        await _db.SaveChangesAsync();
        return created;
    }

    private static TaskArea InferArea(string source, string _)
    {
        var lowered = source.ToLowerInvariant();
        if (lowered.Contains("svp") || lowered.Contains("spravna vyrobni praxe") || lowered.Contains("inspekce"))
        {
            return TaskArea.Svp;
        }
        if (lowered.Contains("sdp") || lowered.Contains("spravna distribucni praxe") || lowered.Contains("distribuc"))
        {
            return TaskArea.Sdp;
        }
        if (lowered.Contains("bozp") || lowered.Contains("bezpecnost prace") || lowered.Contains("skoleni"))
        {
            return TaskArea.Bozp;
        }
        if (lowered.Contains("po") || lowered.Contains("pozar") || lowered.Contains("hasic") || lowered.Contains("evakuac"))
        {
            return TaskArea.Po;
        }
        return TaskArea.Jine;
    }
}
