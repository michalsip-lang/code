using AkcniPlan.Models;
using Microsoft.EntityFrameworkCore;

namespace AkcniPlan.Data;

public class AppDbContext(DbContextOptions<AppDbContext> options) : DbContext(options)
{
    public DbSet<TaskItem> Tasks => Set<TaskItem>();
    public DbSet<TaskTag> Tags => Set<TaskTag>();
    public DbSet<TaskItemTag> TaskItemTags => Set<TaskItemTag>();
    public DbSet<TaskDependency> TaskDependencies => Set<TaskDependency>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<TaskItemTag>()
            .HasKey(x => new { x.TaskItemId, x.TaskTagId });

        modelBuilder.Entity<TaskItemTag>()
            .HasOne(x => x.TaskItem)
            .WithMany(x => x.TaskTags)
            .HasForeignKey(x => x.TaskItemId);

        modelBuilder.Entity<TaskItemTag>()
            .HasOne(x => x.TaskTag)
            .WithMany(x => x.TaskItems)
            .HasForeignKey(x => x.TaskTagId);

        modelBuilder.Entity<TaskDependency>()
            .HasKey(x => new { x.TaskItemId, x.DependsOnTaskItemId });

        modelBuilder.Entity<TaskDependency>()
            .HasOne(x => x.TaskItem)
            .WithMany(x => x.DependsOn)
            .HasForeignKey(x => x.TaskItemId)
            .OnDelete(DeleteBehavior.Restrict);

        modelBuilder.Entity<TaskDependency>()
            .HasOne(x => x.DependsOnTaskItem)
            .WithMany(x => x.RequiredBy)
            .HasForeignKey(x => x.DependsOnTaskItemId)
            .OnDelete(DeleteBehavior.Restrict);

        modelBuilder.Entity<TaskTag>()
            .HasIndex(x => x.Name)
            .IsUnique();
    }
}
