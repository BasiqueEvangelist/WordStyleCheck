using System.Text.Json;
using Microsoft.AspNetCore.Mvc;
using WordStyleCheckService.Worker;

namespace WordStyleCheckService.Foreman.Controllers;

[ApiController]
public class TaskController(Db db) : ControllerBase
{
    [HttpPost]
    [Route("/upload-file")]
    public async Task<string> UploadFile([FromBody] string url)
    {
        var result = await db.EnqueueTask("LintFile", JsonSerializer.Serialize(new TaskInputs()
        {
            DownloadUrl = url
        }));
        
        return result.ToString();
    }

    [HttpGet]
    [Route("/task/{id}")]
    public async Task<IActionResult> TaskDetails(uint id)
    {
        var info = await db.GetTask(id);
        if (info == null) return NotFound();

        if (info.ResultData == null)
        {
            return Ok(new
            {
                Code = "pending"
            });
        }

        return Ok(info.ResultData);
    }
}
