using AkcniPlan.Models;

namespace AkcniPlan.Services;

public class PriorityScoringService
{
    public int CalculateScore(TaskItem task, DateTime utcNow)
    {
        if (task.Status == TaskItemStatus.Done)
        {
            return 0;
        }

        var dueDateScore = 0.0;
        var overdueScore = 0.0;

        if (task.DueDate.HasValue)
        {
            var daysToDue = (task.DueDate.Value.Date - utcNow.Date).TotalDays;
            if (daysToDue >= 0)
            {
                dueDateScore = 100 - Math.Min(100, daysToDue * 12.5);
            }
            else
            {
                dueDateScore = 100;
                overdueScore = Math.Min(100, Math.Abs(daysToDue) * 20);
            }
        }

        var effortScore = Math.Clamp((task.EstimatedHours / 16.0) * 100.0, 0, 100);
        var importanceScore = Math.Clamp((task.Importance / 5.0) * 100.0, 0, 100);
        var dependencyScore = Math.Clamp(task.DependsOn.Count * 15.0, 0, 100);

        var weighted =
            dueDateScore * 0.35 +
            effortScore * 0.15 +
            importanceScore * 0.30 +
            overdueScore * 0.10 +
            dependencyScore * 0.10;

        return (int)Math.Round(Math.Clamp(weighted, 0, 100));
    }
}
