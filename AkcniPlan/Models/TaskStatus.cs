using System.ComponentModel.DataAnnotations;

namespace AkcniPlan.Models;

public enum TaskItemStatus
{
    [Display(Name = "K vyřízení")]
    Todo = 0,

    [Display(Name = "Rozpracováno")]
    InProgress = 1,

    [Display(Name = "Hotovo")]
    Done = 2,

    [Display(Name = "Blokováno")]
    Blocked = 3
}
