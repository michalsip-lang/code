using AkcniPlan.Data;
using AkcniPlan.Models;
using AkcniPlan.Services;
using AkcniPlan.ViewModels;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace AkcniPlan.Controllers;

public class TasksController(
    AppDbContext db,
    PriorityScoringService scoring,
    DashboardService dashboardService,
    TaskAutomationService automationService) : Controller
{
    public async Task<IActionResult> Index(string? filter = null)
    {
        await dashboardService.RecalculatePrioritiesAsync();

        var today = DateTime.UtcNow.Date;
        var weekEnd = today.AddDays(7);

        IQueryable<TaskItem> query = db.Tasks;
        string? filterLabel = null;

        switch ((filter ?? string.Empty).Trim().ToLowerInvariant())
        {
            case "overdue":
                query = query.Where(t => t.Status != TaskItemStatus.Done && t.DueDate != null && t.DueDate < today);
                filterLabel = "Po termínu";
                break;
            case "today":
                query = query.Where(t => t.Status != TaskItemStatus.Done && t.DueDate != null && t.DueDate.Value.Date == today);
                filterLabel = "Na dnešek";
                break;
            case "week":
                query = query.Where(t => t.Status != TaskItemStatus.Done && t.DueDate != null && t.DueDate >= today && t.DueDate <= weekEnd);
                filterLabel = "Tento týden";
                break;
            case "completed":
                query = query.Where(t => t.Status == TaskItemStatus.Done);
                filterLabel = "Dokončené";
                break;
            case "inprogress":
                query = query.Where(t => t.Status == TaskItemStatus.InProgress);
                filterLabel = "Rozpracované";
                break;
        }

        var tasks = await query
            .Include(t => t.TaskTags)
            .ThenInclude(tt => tt.TaskTag)
            .OrderByDescending(t => t.PriorityScore)
            .ThenBy(t => t.DueDate)
            .ToListAsync();

        ViewBag.AllTasks = await db.Tasks.OrderBy(t => t.Title).ToListAsync();
        ViewBag.FilterLabel = filterLabel;
        return View(tasks);
    }

    public async Task<IActionResult> Create()
    {
        ViewBag.AllTasks = await db.Tasks.OrderBy(t => t.Title).ToListAsync();
        return View(new TaskFormViewModel());
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(TaskFormViewModel model)
    {
        if (!ModelState.IsValid)
        {
            ViewBag.AllTasks = await db.Tasks.OrderBy(t => t.Title).ToListAsync();
            return View(model);
        }

        var task = new TaskItem
        {
            Title = model.Title,
            Description = model.Description,
            DueDate = model.DueDate,
            EstimatedHours = model.EstimatedHours,
            ActualHours = model.ActualHours,
            Importance = model.Importance,
            Area = model.Area,
            Status = model.Status,
            CreatedAt = DateTime.UtcNow,
        };

        task.PriorityScore = scoring.CalculateScore(task, DateTime.UtcNow);
        db.Tasks.Add(task);
        await db.SaveChangesAsync();

        await UpdateTagsAsync(task.Id, model.TagsCsv);
        await UpdateDependenciesAsync(task.Id, model.DependencyIds);
        TempData["StatusMessage"] = "Úkol byl uložen.";
        return RedirectToAction(nameof(Index));
    }

    public async Task<IActionResult> Edit(int id)
    {
        var task = await db.Tasks
            .Include(t => t.TaskTags)
            .ThenInclude(tt => tt.TaskTag)
            .Include(t => t.DependsOn)
            .FirstOrDefaultAsync(t => t.Id == id);

        if (task is null)
        {
            return NotFound();
        }

        var vm = new TaskFormViewModel
        {
            Id = task.Id,
            Title = task.Title,
            Description = task.Description,
            DueDate = task.DueDate,
            EstimatedHours = task.EstimatedHours,
            ActualHours = task.ActualHours,
            Importance = task.Importance,
            Area = task.Area,
            Status = task.Status,
            TagsCsv = string.Join(", ", task.TaskTags.Select(x => x.TaskTag.Name)),
            DependencyIds = task.DependsOn.Select(x => x.DependsOnTaskItemId).ToList()
        };

        ViewBag.AllTasks = await db.Tasks.Where(t => t.Id != id).OrderBy(t => t.Title).ToListAsync();
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(TaskFormViewModel model)
    {
        if (!ModelState.IsValid || model.Id is null)
        {
            ViewBag.AllTasks = await db.Tasks.Where(t => t.Id != model.Id).OrderBy(t => t.Title).ToListAsync();
            return View(model);
        }

        var task = await db.Tasks.Include(t => t.DependsOn).FirstOrDefaultAsync(t => t.Id == model.Id.Value);
        if (task is null)
        {
            return NotFound();
        }

        task.Title = model.Title;
        task.Description = model.Description;
        task.DueDate = model.DueDate;
        task.EstimatedHours = model.EstimatedHours;
        task.ActualHours = model.ActualHours;
        task.Importance = model.Importance;
        task.Area = model.Area;
        task.Status = model.Status;

        if (task.Status == TaskItemStatus.Done && task.CompletedAt is null)
        {
            task.CompletedAt = DateTime.UtcNow;
        }
        if (task.Status != TaskItemStatus.Done)
        {
            task.CompletedAt = null;
        }

        task.PriorityScore = scoring.CalculateScore(task, DateTime.UtcNow);
        await db.SaveChangesAsync();

        await UpdateTagsAsync(task.Id, model.TagsCsv);
        await UpdateDependenciesAsync(task.Id, model.DependencyIds);

        TempData["StatusMessage"] = "Změny úkolu byly uloženy.";
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ToggleDone(int id)
    {
        var task = await db.Tasks.Include(t => t.DependsOn).FirstOrDefaultAsync(t => t.Id == id);
        if (task is null)
        {
            return NotFound();
        }

        if (task.Status == TaskItemStatus.Done)
        {
            task.Status = TaskItemStatus.Todo;
            task.CompletedAt = null;
        }
        else
        {
            task.Status = TaskItemStatus.Done;
            task.CompletedAt = DateTime.UtcNow;
            task.ActualHours = Math.Max(task.ActualHours, task.EstimatedHours);
        }

        task.PriorityScore = scoring.CalculateScore(task, DateTime.UtcNow);
        await db.SaveChangesAsync();
        TempData["StatusMessage"] = task.Status == TaskItemStatus.Done
            ? "Úkol byl označen jako hotový."
            : "Úkol byl vrácen mezi otevřené.";
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Delete(int id)
    {
        var task = await db.Tasks.FirstOrDefaultAsync(t => t.Id == id);
        if (task is not null)
        {
            db.Tasks.Remove(task);
            await db.SaveChangesAsync();
        }
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> AutoGenerate(string autoInput)
    {
        var count = await automationService.GenerateFromTextAsync(autoInput);
        TempData["StatusMessage"] = count == 0
            ? "Automatické vytvoření nenašlo žádné nové úkoly."
            : $"Automaticky vytvořeno úkolů: {count}.";
        return RedirectToAction(nameof(Index));
    }

    private async Task UpdateTagsAsync(int taskId, string csv)
    {
        var existing = await db.TaskItemTags.Where(x => x.TaskItemId == taskId).ToListAsync();
        if (existing.Count > 0)
        {
            db.TaskItemTags.RemoveRange(existing);
            await db.SaveChangesAsync();
        }

        var tags = csv.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        foreach (var tagName in tags)
        {
            var tag = await db.Tags.FirstOrDefaultAsync(t => t.Name.ToLower() == tagName.ToLower());
            if (tag is null)
            {
                tag = new TaskTag { Name = tagName };
                db.Tags.Add(tag);
                await db.SaveChangesAsync();
            }

            db.TaskItemTags.Add(new TaskItemTag { TaskItemId = taskId, TaskTagId = tag.Id });
        }

        await db.SaveChangesAsync();
    }

    private async Task UpdateDependenciesAsync(int taskId, List<int> dependencyIds)
    {
        var existing = await db.TaskDependencies.Where(x => x.TaskItemId == taskId).ToListAsync();
        if (existing.Count > 0)
        {
            db.TaskDependencies.RemoveRange(existing);
            await db.SaveChangesAsync();
        }

        var clean = dependencyIds.Where(x => x != taskId).Distinct().ToList();
        foreach (var depId in clean)
        {
            db.TaskDependencies.Add(new TaskDependency
            {
                TaskItemId = taskId,
                DependsOnTaskItemId = depId
            });
        }

        await db.SaveChangesAsync();
    }
}
