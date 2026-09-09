using System.ComponentModel.DataAnnotations;
using AkcniPlan.Models;

namespace AkcniPlan.ViewModels;

public class TaskFormViewModel
{
    public int? Id { get; set; }

    [Required]
    [MaxLength(160)]
    public string Title { get; set; } = string.Empty;

    [MaxLength(4000)]
    public string Description { get; set; } = string.Empty;

    public DateTime? DueDate { get; set; }

    [Range(0.25, 1000)]
    public double EstimatedHours { get; set; } = 1;

    [Range(0, 1000)]
    public double ActualHours { get; set; } = 0;

    [Range(1, 5)]
    public int Importance { get; set; } = 3;

    public TaskArea Area { get; set; } = TaskArea.Jine;

    public TaskItemStatus Status { get; set; } = TaskItemStatus.Todo;

    public string TagsCsv { get; set; } = string.Empty;

    public List<int> DependencyIds { get; set; } = [];

    public string AutoInput { get; set; } = string.Empty;
}
