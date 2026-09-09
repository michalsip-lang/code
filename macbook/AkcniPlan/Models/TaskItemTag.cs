namespace AkcniPlan.Models;

public class TaskItemTag
{
    public int TaskItemId { get; set; }
    public TaskItem TaskItem { get; set; } = null!;

    public int TaskTagId { get; set; }
    public TaskTag TaskTag { get; set; } = null!;
}
