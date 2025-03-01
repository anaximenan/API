using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.OpenApi.Models;
using PdfApi.Filters;

var builder = WebApplication.CreateBuilder(args);

// Configuración de servicios
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();

// Configuración de Swagger
builder.Services.AddSwaggerGen(c =>
{
    c.SwaggerDoc("v1", new OpenApiInfo 
    { 
        Title = "PDF API",
        Version = "v1",
        Description = "API para procesamiento de PDFs bancarios",
        Contact = new OpenApiContact
        {
            Name = "Soporte Técnico",
            Email = "soporte@pdfapi.com"
        }
    });

    c.OperationFilter<FileUploadOperationFilter>();
});

// Configuración CORS
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll", policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyMethod()
              .AllowAnyHeader();
    });
});

var app = builder.Build();

// 🚀 Mostrar Swagger SIEMPRE (Render no tiene modo "Development")
app.UseSwagger();
app.UseSwaggerUI(c =>
{
    c.SwaggerEndpoint("/swagger/v1/swagger.json", "PDF API v1");
    c.RoutePrefix = "swagger"; // Ruta de acceso
});

// ❌ 🔴 NO redirigir a HTTPS en Render
// app.UseHttpsRedirection();

app.UseRouting();
app.UseCors("AllowAll");
app.UseAuthorization();

app.MapControllers();

// 📌 🔥 Forzar puerto correcto en Render
var port = Environment.GetEnvironmentVariable("PORT") ?? "8080";
app.Run($"http://0.0.0.0:{port}");
