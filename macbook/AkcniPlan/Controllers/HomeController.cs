using AkcniPlan.Services;
using Microsoft.AspNetCore.Mvc;

namespace AkcniPlan.Controllers;

public class HomeController(DashboardService dashboardService) : Controller
{
    public async Task<IActionResult> Index()
    {
        var model = await dashboardService.BuildAsync();
        return View(model);
    }

    public IActionResult Error()
    {
        return View();
    }
}
