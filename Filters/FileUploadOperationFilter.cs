using Microsoft.AspNetCore.Mvc.Controllers;
using Microsoft.OpenApi.Models;
using Swashbuckle.AspNetCore.SwaggerGen;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using Microsoft.AspNetCore.Http;
using System.Reflection;

namespace PdfApi.Filters
{
    public class FileUploadOperationFilter : IOperationFilter
    {
        public void Apply(OpenApiOperation operation, OperationFilterContext context)
        {
            var actionDescriptor = context.ApiDescription.ActionDescriptor as ControllerActionDescriptor;
            if (actionDescriptor == null) return;

            var formParameters = actionDescriptor.Parameters
                .Where(p => p.BindingInfo?.BindingSource?.Id == "Form")
                .OfType<ControllerParameterDescriptor>()
                .ToList();

            if (!formParameters.Any()) return;

            var properties = new Dictionary<string, OpenApiSchema>();
            var required = new HashSet<string>();

            foreach (var param in formParameters)
            {
                var isRequired = param.ParameterInfo
                    .GetCustomAttributes(typeof(RequiredAttribute), false)
                    .Any();

                var schema = param.ParameterType switch
                {
                    _ when param.ParameterType == typeof(IFormFile) => new OpenApiSchema { Type = "string", Format = "binary" },
                    _ when param.ParameterType == typeof(IFormFile[]) => new OpenApiSchema { Type = "array", Items = new OpenApiSchema { Type = "string", Format = "binary" } },
                    _ => new OpenApiSchema { Type = "integer" }
                };

                properties.Add(param.Name, schema);
                if (isRequired) required.Add(param.Name);
            }

            operation.RequestBody = new OpenApiRequestBody
            {
                Content = { ["multipart/form-data"] = new OpenApiMediaType { Schema = new OpenApiSchema {
                    Type = "object",
                    Properties = properties,
                    Required = required
                }}}
            };

            // Eliminar parámetros duplicados (CORREGIDO)
            var parametersToRemove = operation.Parameters
                .Where(p => formParameters.Any(fp => fp.Name == p.Name))
                .ToList();

            foreach (var param in parametersToRemove)
            {
                operation.Parameters.Remove(param);
            }
        }
    }
}