using System.ComponentModel.DataAnnotations;

namespace AkcniPlan.Models;

public class TaskItem
{
    public int Id { get; set; }

    [Required]
    [MaxLength(160)]
    public string Title { get; set; } = string.Empty;

    [MaxLength(4000)]
    public string Description { get; set; } = string.Empty;

    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

    public DateTime? DueDate { get; set; }

    [Range(0.25, 1000)]
    public double EstimatedHours { get; set; } = 1;

    [Range(0, 1000)]
    public double ActualHours { get; set; } = 0;

    [Range(1, 5)]
    public int Importance { get; set; } = 3;

    public TaskArea Area { get; set; } = TaskArea.Jine;

    public TaskItemStatus Status { get; set; } = TaskItemStatus.Todo;

    [Range(0, 100)]
    public int PriorityScore { get; set; } = 0;

    public DateTime? CompletedAt { get; set; }

    [MaxLength(512)]
    public string? SourceMessageId { get; set; }

    public List<TaskItemTag> TaskTags { get; set; } = [];

    public List<TaskDependency> DependsOn { get; set; } = [];

    public List<TaskDependency> RequiredBy { get; set; } = [];
}
