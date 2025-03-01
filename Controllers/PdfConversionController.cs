using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;

namespace PdfApi.Controllers
{
    [ApiController]
    [Route("[controller]")]
    [IgnoreAntiforgeryToken] // Se deshabilita la validación antiforgery para todos los endpoints de este controller
    public class ConvertController : ControllerBase
    {
        // Este endpoint se encargará de recibir el archivo PDF y simular su conversión a Excel.
        [HttpPost]
        public async Task<IActionResult> ConvertPdfToExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No se proporcionó un archivo válido.");
            }

            // Guardar el archivo PDF en una ruta temporal
            var tempPath = Path.GetTempPath();
            var pdfFilePath = Path.Combine(tempPath, file.FileName);

            using (var stream = new FileStream(pdfFilePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            // Aquí se debería implementar la lógica real de conversión.
            // Por ahora, se simula la conversión cambiando la extensión del archivo.
            string excelFilePath = pdfFilePath.Replace(".pdf", ".xlsx");

            return Ok(new { message = "Archivo recibido correctamente", excelFilePath });
        }
    }
}
