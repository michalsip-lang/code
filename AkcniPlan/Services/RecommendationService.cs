using AkcniPlan.Models;

namespace AkcniPlan.Services;

public class RecommendationService
{
    public List<string> BuildRecommendations(IEnumerable<TaskItem> tasks, double dailyCapacityHours)
    {
        var list = tasks.ToList();
        var now = DateTime.UtcNow;
        var messages = new List<string>();

        var top = list
            .Where(t => t.Status != TaskItemStatus.Done)
            .OrderByDescending(t => t.PriorityScore)
            .ThenBy(t => t.DueDate)
            .Take(3)
            .ToList();

        if (top.Count > 0)
        {
            messages.Add($"Řešit jako první: {string.Join(", ", top.Select(t => t.Title))}.");
        }

        var next3DaysLoad = list
            .Where(t => t.Status != TaskItemStatus.Done && t.DueDate != null && t.DueDate <= now.Date.AddDays(3))
            .Sum(t => t.EstimatedHours);

        if (next3DaysLoad > dailyCapacityHours * 3)
        {
            messages.Add("V následujících 3 dnech hrozí přetížení. Přesuňte část úkolů nebo snižte scope.");
        }

        messages.Add($"Realistická denní kapacita: {dailyCapacityHours:F1} h.");

        var remaining = list
            .Where(t => t.Status != TaskItemStatus.Done)
            .Sum(t => Math.Max(0, t.EstimatedHours - t.ActualHours));

        var days = dailyCapacityHours <= 0 ? 0 : Math.Ceiling(remaining / dailyCapacityHours);
        if (days > 0)
        {
            messages.Add($"Odhad dokončení aktuálního backlogu: {now.Date.AddDays((int)days):yyyy-MM-dd}.");
        }

        return messages;
    }
}
