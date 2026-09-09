using AkcniPlan.Data;
using AkcniPlan.Models;
using AkcniPlan.ViewModels;
using Microsoft.EntityFrameworkCore;

namespace AkcniPlan.Services;

public class DashboardService(
    AppDbContext db,
    PriorityScoringService scoring,
    AnalyticsService analytics,
    RecommendationService recommendations)
{
    public async Task RecalculatePrioritiesAsync()
    {
        var tasks = await db.Tasks
            .Include(t => t.DependsOn)
            .Where(t => t.Status != TaskItemStatus.Done)
            .ToListAsync();

        foreach (var task in tasks)
        {
            task.PriorityScore = scoring.CalculateScore(task, DateTime.UtcNow);
        }

        await db.SaveChangesAsync();
    }

    public async Task<DashboardViewModel> BuildAsync()
    {
        await RecalculatePrioritiesAsync();

        var now = DateTime.UtcNow;
        var weekEnd = now.Date.AddDays(7);

        var tasks = await db.Tasks
            .Include(t => t.DependsOn)
            .OrderByDescending(t => t.PriorityScore)
            .ToListAsync();

        var openTasks = tasks.Where(t => t.Status != TaskItemStatus.Done).ToList();
        var completed = tasks.Where(t => t.Status == TaskItemStatus.Done).ToList();

        var completionRate = tasks.Count == 0 ? 0 : completed.Count * 100.0 / tasks.Count;

        var avgCompletion = await analytics.AverageCompletionHoursAsync();
        var onTimeRate = await analytics.OnTimeRateAsync();
        var productivity = await analytics.ProductivityIndexAsync();

        var dailyCapacity = await EstimateDailyCapacityAsync();

        return new DashboardViewModel
        {
            OverdueCount = openTasks.Count(t => t.DueDate != null && t.DueDate < now.Date),
            TodayCount = openTasks.Count(t => t.DueDate?.Date == now.Date),
            ThisWeekCount = openTasks.Count(t => t.DueDate != null && t.DueDate >= now.Date && t.DueDate <= weekEnd),
            CompletedCount = completed.Count,
            InProgressCount = tasks.Count(t => t.Status == TaskItemStatus.InProgress),
            CompletionRate = Math.Round(completionRate, 2),
            OnTimeRate = Math.Round(onTimeRate, 2),
            AverageCompletionHours = Math.Round(avgCompletion, 2),
            ProductivityIndex = productivity,
            TopPriorityTasks = openTasks.Take(10).ToList(),
            Recommendations = recommendations.BuildRecommendations(openTasks, dailyCapacity)
        };
    }

    public async Task<double> EstimateDailyCapacityAsync()
    {
        var from = DateTime.UtcNow.Date.AddDays(-14);
        var done = await db.Tasks
            .Where(t => t.Status == TaskItemStatus.Done && t.CompletedAt != null && t.CompletedAt >= from)
            .ToListAsync();

        if (done.Count == 0)
        {
            return 4.0;
        }

        var daysWithWork = done
            .GroupBy(t => t.CompletedAt!.Value.Date)
            .Select(g => g.Sum(x => x.ActualHours > 0 ? x.ActualHours : x.EstimatedHours))
            .ToList();

        if (daysWithWork.Count == 0)
        {
            return 4.0;
        }

        var avg = daysWithWork.Average();
        return Math.Round(Math.Clamp(avg, 2.0, 10.0), 1);
    }
}
