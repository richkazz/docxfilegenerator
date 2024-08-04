using DocxFileGenerator.Models;
using DoxcFileGenerator.Utils;
using Microsoft.AspNetCore.Mvc;

var builder = WebApplication.CreateBuilder(args);

var port = Environment.GetEnvironmentVariable("PORT") ?? "3000";
var url = $"http://0.0.0.0:{port}";
var target = Environment.GetEnvironmentVariable("TARGET") ?? "World";

var app = builder.Build();

app.MapGet("/", () => $"Hello {target}!");


app.MapPost("/api/resume", async ([FromBody] Resume resume, HttpContext httpContext) =>
{
    try
    {
        // Generate a unique file name
        string fileName = $"resume_{Guid.NewGuid()}.docx";
        string filePath = Path.Combine(Path.GetTempPath(), fileName);

        // Generate the Word document
        ResumeToWordDocumentTemplateTwo.GenerateWordDocument(resume, filePath);

        // Read the file into memory
        byte[] fileBytes = await File.ReadAllBytesAsync(filePath);

        // Delete the temporary file
        File.Delete(filePath);

        // Set the content type and attachment header
        httpContext.Response.ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        httpContext.Response.Headers.Append("Content-Disposition", $"attachment; filename={fileName}");

        // Return the file
        await httpContext.Response.Body.WriteAsync(fileBytes);
    }
    catch (Exception ex)
    {
        // Log the exception
        Console.WriteLine($"Error generating resume: {ex.Message}");
        // Return an error response
        await httpContext.Response.WriteAsync($"Error generating resume: {ex.Message}");
    }
});

app.Run(url);