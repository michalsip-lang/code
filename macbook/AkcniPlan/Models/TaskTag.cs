using System.ComponentModel.DataAnnotations;

namespace AkcniPlan.Models;

public class TaskTag
{
    public int Id { get; set; }

    [Required]
    [MaxLength(50)]
    public string Name { get; set; } = string.Empty;

    public List<TaskItemTag> TaskItems { get; set; } = [];
}
