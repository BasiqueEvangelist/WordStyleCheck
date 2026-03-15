using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using Minio;
using Minio.DataModel.Args;
using WordStyleCheckService.Worker;
using Options = WordStyleCheckService.Worker.Options;

namespace WordStyleCheckService.Foreman.Controllers;

[ApiController]
public class TaskController(Db db, IMinioClient s3, IOptionsSnapshot<Options> options) : ControllerBase
{
    [HttpPost]
    [Route("/upload-file")]
    // [RequestSizeLimit(512 * 1024 * 1024)]
    // [RequestFormLimits(MultipartBodyLengthLimit = 512 * 1024 * 1024)]

    public async Task<string> UploadFile([FromForm] IFormFile file)
    {
        string objectKey = RandomNumberGenerator.GetHexString(32) + ".docx";

        MemoryStream ms = new MemoryStream();
        await file.CopyToAsync(ms);
        ms.Seek(0, SeekOrigin.Begin);
        
        await s3.PutObjectAsync(new PutObjectArgs()
            .WithBucket(options.Value.S3IngressBucket)
            .WithObject(objectKey)
            .WithStreamData(ms)
            .WithObjectSize(ms.Length)
            .WithContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
        
        var result = await db.EnqueueTask("LintFile", JsonSerializer.Serialize(new TaskInputs()
        {
            InputObject = objectKey
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
    
    [HttpGet]
    [Route("/task/{id}/download")]
    public async Task<IActionResult> TaskDownload(uint id)
    {
        var info = await db.GetTask(id);
        if (info == null) return NotFound();

        if (info.ResultData == null)
        {
            return NotFound();
        }

        var outp = JsonSerializer.Deserialize<TaskOutputs>(info.ResultData)!;

        if (outp.OutputObject == null) return NotFound();

        var ms = new MemoryStream();
        await s3.GetObjectAsync(new GetObjectArgs()
            .WithBucket(options.Value.S3EgressBucket)
            .WithObject(outp.OutputObject)
            .WithCallbackStream((s, token) => s.CopyToAsync(ms, token)));
        ms.Seek(0, SeekOrigin.Begin);

        return File(ms, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    }
}
