using AkcniPlan.Models;

namespace AkcniPlan.ViewModels;

public class DashboardViewModel
{
    public int OverdueCount { get; set; }
    public int TodayCount { get; set; }
    public int ThisWeekCount { get; set; }
    public int CompletedCount { get; set; }
    public int InProgressCount { get; set; }

    public double CompletionRate { get; set; }
    public double OnTimeRate { get; set; }
    public double AverageCompletionHours { get; set; }
    public double ProductivityIndex { get; set; }

    public List<TaskItem> TopPriorityTasks { get; set; } = [];
    public List<string> Recommendations { get; set; } = [];
}
