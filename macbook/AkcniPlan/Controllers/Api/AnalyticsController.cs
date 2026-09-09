using AkcniPlan.Data;
using AkcniPlan.Models;
using AkcniPlan.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace AkcniPlan.Controllers.Api;

[ApiController]
[Route("api/[controller]")]
public class AnalyticsController(
    AnalyticsService analytics,
    DashboardService dashboard,
    AppDbContext db) : ControllerBase
{
    [HttpGet("series")]
    public async Task<IActionResult> Series()
    {
        var day = await analytics.CompletedPerDayAsync(30);
        var week = await analytics.CompletedPerWeekAsync(12);
        var month = await analytics.CompletedPerMonthAsync(12);

        var statesRaw = await db.Tasks
            .GroupBy(t => t.Status)
            .Select(g => new { Status = g.Key, Count = g.Count() })
            .ToListAsync();

        var states = statesRaw.Select(s => new
        {
            status = s.Status switch
            {
                TaskItemStatus.Todo => "K vyřízení",
                TaskItemStatus.InProgress => "Rozpracováno",
                TaskItemStatus.Done => "Hotovo",
                TaskItemStatus.Blocked => "Blokováno",
                _ => s.Status.ToString()
            },
            count = s.Count
        });

        var heatmap = await analytics.CompletedPerDayAsync(120);

        return Ok(new
        {
            daily = day,
            weekly = week,
            monthly = month,
            states,
            heatmap
        });
    }

    [HttpGet("kpi")]
    public async Task<IActionResult> Kpi()
    {
        var completion = await db.Tasks.CountAsync() == 0
            ? 0
            : await db.Tasks.CountAsync(t => t.Status == TaskItemStatus.Done) * 100.0 / await db.Tasks.CountAsync();

        return Ok(new
        {
            completionRate = Math.Round(completion, 2),
            onTimeRate = Math.Round(await analytics.OnTimeRateAsync(), 2),
            avgCompletionHours = Math.Round(await analytics.AverageCompletionHoursAsync(), 2),
            productivity = await analytics.ProductivityIndexAsync(),
            dailyCapacity = await dashboard.EstimateDailyCapacityAsync()
        });
    }
}
