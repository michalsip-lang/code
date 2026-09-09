namespace AkcniPlan.Models;

public class TaskDependency
{
    public int TaskItemId { get; set; }
    public TaskItem TaskItem { get; set; } = null!;

    public int DependsOnTaskItemId { get; set; }
    public TaskItem DependsOnTaskItem { get; set; } = null!;
}
