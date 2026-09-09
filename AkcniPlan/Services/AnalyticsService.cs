using AkcniPlan.Data;
using AkcniPlan.Models;
using Microsoft.EntityFrameworkCore;

namespace AkcniPlan.Services;

public class AnalyticsService(AppDbContext db)
{
    public async Task<Dictionary<string, int>> CompletedPerDayAsync(int days = 30)
    {
        var from = DateTime.UtcNow.Date.AddDays(-days + 1);
        var tasks = await db.Tasks
            .Where(t => t.Status == TaskItemStatus.Done && t.CompletedAt != null && t.CompletedAt >= from)
            .ToListAsync();

        var series = Enumerable.Range(0, days)
            .Select(i => from.AddDays(i).ToString("yyyy-MM-dd"))
            .ToDictionary(d => d, _ => 0);

        foreach (var task in tasks)
        {
            var key = task.CompletedAt!.Value.Date.ToString("yyyy-MM-dd");
            if (series.ContainsKey(key))
            {
                series[key] += 1;
            }
        }

        return series;
    }

    public async Task<Dictionary<string, int>> CompletedPerMonthAsync(int months = 12)
    {
        var from = new DateTime(DateTime.UtcNow.Year, DateTime.UtcNow.Month, 1).AddMonths(-months + 1);
        var tasks = await db.Tasks
            .Where(t => t.Status == TaskItemStatus.Done && t.CompletedAt != null && t.CompletedAt >= from)
            .ToListAsync();

        var series = Enumerable.Range(0, months)
            .Select(i => from.AddMonths(i).ToString("yyyy-MM"))
            .ToDictionary(d => d, _ => 0);

        foreach (var task in tasks)
        {
            var key = task.CompletedAt!.Value.ToString("yyyy-MM");
            if (series.ContainsKey(key))
            {
                series[key] += 1;
            }
        }

        return series;
    }

    public async Task<Dictionary<string, int>> CompletedPerWeekAsync(int weeks = 12)
    {
        var from = DateTime.UtcNow.Date.AddDays(-(7 * weeks) + 1);
        var tasks = await db.Tasks
            .Where(t => t.Status == TaskItemStatus.Done && t.CompletedAt != null && t.CompletedAt >= from)
            .ToListAsync();

        var series = new Dictionary<string, int>();
        for (var i = 0; i < weeks; i++)
        {
            var weekStart = from.AddDays(i * 7);
            var weekEnd = weekStart.AddDays(6);
            series[$"{weekStart:MM/dd}-{weekEnd:MM/dd}"] = 0;
        }

        foreach (var task in tasks)
        {
            var diff = (task.CompletedAt!.Value.Date - from).TotalDays;
            var index = (int)(diff / 7);
            if (index >= 0 && index < weeks)
            {
                var weekStart = from.AddDays(index * 7);
                var weekEnd = weekStart.AddDays(6);
                var key = $"{weekStart:MM/dd}-{weekEnd:MM/dd}";
                series[key] += 1;
            }
        }

        return series;
    }

    public async Task<double> AverageCompletionHoursAsync()
    {
        var tasks = await db.Tasks
            .Where(t => t.Status == TaskItemStatus.Done && t.CompletedAt != null)
            .ToListAsync();

        if (tasks.Count == 0)
        {
            return 0;
        }

        return tasks.Average(t => (t.CompletedAt!.Value - t.CreatedAt).TotalHours);
    }

    public async Task<double> OnTimeRateAsync()
    {
        var tasks = await db.Tasks
            .Where(t => t.Status == TaskItemStatus.Done)
            .ToListAsync();

        if (tasks.Count == 0)
        {
            return 0;
        }

        var onTime = tasks.Count(t => t.DueDate == null || t.CompletedAt <= t.DueDate);
        return (double)onTime / tasks.Count * 100;
    }

    public async Task<double> ProductivityIndexAsync()
    {
        var from = DateTime.UtcNow.Date.AddDays(-14);
        var completed = await db.Tasks
            .Where(t => t.Status == TaskItemStatus.Done && t.CompletedAt != null && t.CompletedAt >= from)
            .ToListAsync();

        if (completed.Count == 0)
        {
            return 0;
        }

        var weightedSum = completed.Sum(t => t.Importance * Math.Max(0.5, t.EstimatedHours));
        var perDay = weightedSum / 14.0;
        return Math.Round(perDay, 2);
    }
}
