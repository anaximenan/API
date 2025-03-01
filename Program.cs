using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.OpenApi.Models;
using PdfApi.Filters;

var builder = WebApplication.CreateBuilder(args);

// Configuración de servicios
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();

// Configuración de Swagger (versión corregida)
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
    
    c.OperationFilter<FileUploadOperationFilter>(); // Solo un filtro
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

// Configuración del middleware
if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
    app.UseSwagger();
    app.UseSwaggerUI(c =>
    {
        c.SwaggerEndpoint("/swagger/v1/swagger.json", "PDF API v1");
        c.RoutePrefix = "swagger"; // Ruta de acceso
    });
}

app.UseHttpsRedirection();
app.UseRouting();
app.UseCors("AllowAll");
app.UseAuthorization(); // Requiere el servicio de autorización

app.MapControllers();

app.Run();