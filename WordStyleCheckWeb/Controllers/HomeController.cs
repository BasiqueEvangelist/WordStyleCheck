using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using WordStyleCheckWeb.Models;

namespace WordStyleCheckWeb.Controllers;

public class HomeController : Controller
{
    private readonly DocumentProcessingService _processingService;

    public HomeController(DocumentProcessingService processingService)
    {
        _processingService = processingService;
    }
    
    [Route("/")]
    public IActionResult Index()
    {
        return View();
    }

    [HttpPost]
    [Route("/submit-files")]
    public async Task<IActionResult> SubmitFiles([FromForm] List<IFormFile> inputFile)
    {
        var tasks = await Task.WhenAll(inputFile.Select(async x =>
        {
            string tempPath = Path.GetTempFileName();

            using (var w = System.IO.File.OpenWrite(tempPath))
                await x.CopyToAsync(w);
            
            return _processingService.StartTask(x.FileName, tempPath);
        }));

        return View(new TaskListModel()
        {
            Tasks = tasks.ToList()
        });
    }

    [Route("/task/{id}")]
    public IActionResult TaskStatus(Guid id)
    {
        var task = _processingService.GetTask(id);

        if (task == null) return NotFound();

        return View(task);
    }
    
    [Route("/task/{id}/download")]
    public IActionResult TaskDownload(Guid id)
    {
        var task = _processingService.GetTask(id);

        if (task == null) return NotFound();

        string name = Path.GetFileNameWithoutExtension(task.Name) + "-ANNOTATED.docx";

        return File(System.IO.File.OpenRead(task.GetPath()), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", name);
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}
