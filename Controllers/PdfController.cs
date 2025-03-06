using ClosedXML.Excel;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace PdfApi.Controllers
{
    [ApiController]
    [Route("api/pdf")]
    [IgnoreAntiforgeryToken]
    public class PdfController : ControllerBase
    {
        // -----------------------------------------------------------------
        // Endpoint para procesar PDFs tipo BBVA (Actualizado iText7)
        // -----------------------------------------------------------------
        [HttpPost("bbva")]
        [ProducesResponseType(typeof(List<MovimientoBBVA>), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public IActionResult ProcesarBBVA(
            [FromForm][Required] IFormFile file,
            [FromForm][Required] int anio)
        {
            if (file.Length == 0)
                return BadRequest("No se proporcionó un archivo PDF válido.");

            List<MovimientoBBVA> movimientos = new();

            try
            {
                using (var memoryStream = new MemoryStream())
                {
                    file.CopyTo(memoryStream);
                    memoryStream.Position = 0;

                    using (var reader = new PdfReader(memoryStream))
                    using (var pdfDoc = new PdfDocument(reader))
                    {
                        for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
                        {
                            var strategy = new SimpleTextExtractionStrategy();
                            string pageText = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy);
                            movimientos.AddRange(ExtraerMovimientosBBVA(pageText, anio));
                        }
                    }
                }
                return Ok(movimientos);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al procesar PDF BBVA: {ex.Message}");
            }
        }

        private List<MovimientoBBVA> ExtraerMovimientosBBVA(string pagina, int selectedYear)
        {
            List<MovimientoBBVA> movimientos = new();

            string[] ignoreLines = {
                "Estimado Cliente,",
                "También le informamos que su Contrato ha sido modificado,",
                "Estado de Cuenta Modificado:",
                "Su Estado de Cuenta ha sido modificado y ahora tiene más detalle de información.",
                "Le informamos que su Contrato ha sido modificado, el cual puede consultarlo en cualquier sucursal o en www.bancomer.com",
                "Con Bancomer, adelante,",
                "BBVA Bancomer, S.A.",
                "Institución de Banca Múltiple, Grupo Financiero BBVA Bancomer",
                "Av. Paseo de la Reforma 510, Col. Juárez, Del. Cuauhtémoc, C.P. 06600, Ciudad de México, México,",
                "R.F.C. BBA830831LJ2",
                "el cual puede consultarlo en cualquier sucursal o www.bancomer.com",
                "Con Bancomer, adelante.",
                "BBVA BANCOMER, S.A. INSTITUCION DE BANCA MULTIPLE, GRUPO FINANCIERO BBVA BANCOMER",
                "Total de Movimientos"
            };

            string[] lineas = pagina.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            MovimientoBBVA? currentMovimiento = null;
            bool esReferencia = false;

            for (int i = 0; i < lineas.Length; i++)
            {
                string linea = lineas[i].Trim();
                bool isIgnoredLine = ignoreLines.Any(ign => linea.Contains(ign)) || string.IsNullOrWhiteSpace(linea);
                if (isIgnoredLine) continue;

                var match = Regex.Match(linea, @"^(?<dia1>\d{2})/(?<mes1>[A-Z]{3})\s+(?<dia2>\d{2})/(?<mes2>[A-Z]{3})\s*(?<resto>.*)$");
                if (match.Success)
                {
                    esReferencia = false;
                    string resto = match.Groups["resto"].Value.Trim();

                    // Extraer montos
                    var montosMatch = Regex.Matches(resto, @"\d{1,3}(,\d{3})*(\.\d{2})");
                    string cargosAbonos = "";
                    string operacion = "";
                    string liquidacion = "";

                    if (montosMatch.Count > 0)
                    {
                        cargosAbonos = montosMatch[0].Value;
                        if (montosMatch.Count >= 2) operacion = montosMatch[1].Value;
                        if (montosMatch.Count >= 3) liquidacion = montosMatch[2].Value;

                        foreach (Match monto in montosMatch.Cast<Match>())
                        {
                            resto = resto.Replace(monto.Value, "").Trim();
                        }
                    }

                    // Procesar referencia
                    var refMatch = Regex.Match(resto, @"^(?<descripcion>.*?)(Ref\.\s*)(?<referencia>.*)$");
                    string codDescripcion = resto;
                    string referencia = "";

                    if (refMatch.Success)
                    {
                        codDescripcion = refMatch.Groups["descripcion"].Value.Trim();
                        referencia = "Ref. " + refMatch.Groups["referencia"].Value.Trim();
                    }

                    // Crear movimiento
                    currentMovimiento = new MovimientoBBVA
                    {
                        OPER = $"{match.Groups["dia1"].Value}-{match.Groups["mes1"].Value}",
                        LIQ = $"{match.Groups["dia2"].Value}-{match.Groups["mes2"].Value}",
                        ANIO = selectedYear,
                        COD_DESCRIPCION = codDescripcion,
                        REFERENCIA = referencia,
                        CARGOS_ABONOS = cargosAbonos,
                        OPERACION = operacion,
                        LIQUIDACION = liquidacion
                    };

                    // Procesar líneas siguientes para descripción
                    while (i + 1 < lineas.Length)
                    {
                        string nextLine = lineas[i + 1].Trim();
                        if (ignoreLines.Any(ign => nextLine.Contains(ign))) break;

                        bool tieneMontos = Regex.IsMatch(nextLine, @"\d{1,3}(,\d{3})*(\.\d{2})");
                        bool esNuevoMovimiento = Regex.IsMatch(nextLine, @"^\d{2}/[A-Z]{3}\s+\d{2}/[A-Z]{3}");
                        
                        if (tieneMontos || esNuevoMovimiento) break;

                        currentMovimiento.COD_DESCRIPCION += " " + nextLine;
                        i++;
                    }

                    movimientos.Add(currentMovimiento);
                }
                else if (currentMovimiento != null)
                {
                    if (linea.Contains("Ref."))
                    {
                        esReferencia = true;
                        currentMovimiento.REFERENCIA += $" {linea}";
                    }
                    else if (esReferencia)
                    {
                        currentMovimiento.REFERENCIA += $" {linea}";
                    }
                    else
                    {
                        currentMovimiento.COD_DESCRIPCION += $" {linea}";
                    }
                }
            }

            return movimientos;
        }
    

        // -----------------------------------------------------------------
        // Endpoint para procesar PDFs tipo BanBajío (Actualizado iText7)
        // -----------------------------------------------------------------
        [HttpPost("banbajio")]
        [ProducesResponseType(typeof(List<MovimientoBanBajio>), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public IActionResult ProcesarBanBajio(
            [FromForm][Required] IFormFile file,
            [FromForm][Required] int anio)
        {
            if (file.Length == 0)
                return BadRequest("No se proporcionó un archivo PDF válido.");

            List<MovimientoBanBajio> movimientos = new();

            try
            {
                using (var memoryStream = new MemoryStream())
                {
                    file.CopyTo(memoryStream);
                    memoryStream.Position = 0;

                    using (var reader = new PdfReader(memoryStream))
                    using (var pdfDoc = new PdfDocument(reader))
                    {
                        for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
                        {
                            var strategy = new SimpleTextExtractionStrategy();
                            string pageText = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy);
                            movimientos.AddRange(ExtraerMovimientosBanBajio(pageText, anio));
                        }
                    }
                }
                return Ok(movimientos);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al procesar PDF BanBajío: {ex.Message}");
            }
        }

        private List<MovimientoBanBajio> ExtraerMovimientosBanBajio(string pagina, int selectedYear)
        {
            List<MovimientoBanBajio> movimientos = new();

            string[] ignoreLines = {
                "SALDO ANTERIOR", "SALDO PROMEDIO", "SALDO ACTUAL",
                "TASA ANUAL", "ISR",
                "DETALLE DE LA CUENTA", "DESCRIPCION DE LA OPERACION",
                "FECHA", "NO. REF/DOCT"
            };

            string[] stopPhrases = {
                "SALDO TOTAL",
                "TOTAL DE MOVIMIENTOS EN EL PERIODO"
            };

            bool stopReading = false;
            string[] lineas = pagina.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            MovimientoBanBajio? currentMovimiento = null;

            for (int i = 0; i < lineas.Length; i++)
            {
                if (stopReading) break;

                string linea = lineas[i].Trim();
                if (string.IsNullOrWhiteSpace(linea)) continue;

                if (stopPhrases.Any(sp => linea.Contains(sp, StringComparison.OrdinalIgnoreCase)))
                {
                    stopReading = true;
                    break;
                }
                if (ignoreLines.Any(ign => linea.Contains(ign, StringComparison.OrdinalIgnoreCase))) continue;

                var matchFecha = Regex.Match(linea, @"^(?<dia>\d{1,2})\s+(?<mes>[A-Z]{3})\s+(?<resto>.*)$");
                if (matchFecha.Success)
                {
                    currentMovimiento = new MovimientoBanBajio
                    {
                        FECHA = $"{matchFecha.Groups["dia"].Value}-{matchFecha.Groups["mes"].Value}",
                        ANIO = selectedYear,
                        REF_DOCT = string.Empty,
                        DESCRIPCION = string.Empty,
                        DEPOSITOS_RETIROS = string.Empty,
                        SALDO = string.Empty
                    };

                    string resto = matchFecha.Groups["resto"].Value;
                    var montosMatch = Regex.Matches(resto, @"\$?\s?\d{1,3}(?:,\d{3})*(?:\.\d{2})");

                    if (montosMatch.Count >= 1)
                        currentMovimiento.DEPOSITOS_RETIROS = montosMatch[0].Value.Replace("$", "").Trim();

                    if (montosMatch.Count >= 2)
                        currentMovimiento.SALDO = montosMatch[1].Value.Replace("$", "").Trim();

                    foreach (Match monto in montosMatch)
                    {
                        resto = resto.Replace(monto.Value, "").Trim();
                    }

                    var refDocMatch = Regex.Match(resto, @"^(?<ref>\S+)\s+(?<desc>.*)$");
                    if (refDocMatch.Success)
                    {
                        currentMovimiento.REF_DOCT = refDocMatch.Groups["ref"].Value;
                        currentMovimiento.DESCRIPCION = refDocMatch.Groups["desc"].Value;
                    }
                    else
                    {
                        currentMovimiento.DESCRIPCION = resto;
                    }

                    movimientos.Add(currentMovimiento);
                }
                else if (currentMovimiento != null)
                {
                    currentMovimiento.DESCRIPCION += $" {linea}";
                }
            }
            return movimientos;
        }

        // -----------------------------------------------------------------
        // Endpoint para procesar PDFs tipo Banamex (Actualizado iText7)
        // -----------------------------------------------------------------
        [HttpPost("banamex")]
        [ProducesResponseType(typeof(List<MovimientoBanamex>), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public IActionResult ProcesarBanamex(
            [FromForm][Required] IFormFile file,
            [FromForm][Required] int anio)
        {
            if (file.Length == 0)
                return BadRequest("No se proporcionó un archivo PDF válido.");

            List<MovimientoBanamex> movimientos = new();

            try
            {
                using (var memoryStream = new MemoryStream())
                {
                    file.CopyTo(memoryStream);
                    memoryStream.Position = 0;
                    string transaccionPendiente = string.Empty;

                    using (var reader = new PdfReader(memoryStream))
                    using (var pdfDoc = new PdfDocument(reader))
                    {
                        for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
                        {
                            var strategy = new SimpleTextExtractionStrategy();
                            string pageText = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy);
                            // Aquí se puede registrar pageText en un log para análisis, si se desea.
                            pageText = ProcessTextBanamex(pageText);
                            transaccionPendiente = ExtractMovementsFromPageBanamex(pageText, movimientos, anio, transaccionPendiente);
                        }
                    }

                    if (!string.IsNullOrEmpty(transaccionPendiente))
                    {
                        // Separamos las líneas del bloque pendiente y las procesamos
                        var lines = transaccionPendiente.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        ProcesarBloques(lines, movimientos, anio);
                    }
                }
                return Ok(movimientos);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al procesar PDF Banamex: {ex.Message}");
            }
        }

        private string ProcessTextBanamex(string text)
        {
            // Normaliza las fechas del formato "dd-MES" a "dd MES"
            return Regex.Replace(text, @"(\d{1,2})-([A-Z]{3})", "$1 $2", RegexOptions.IgnoreCase);
        }

        private string ExtractMovementsFromPageBanamex(string pagina, List<MovimientoBanamex> movimientos, int anio, string transaccionPendiente)
        {
            string[] lineas = pagina.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            bool enDetalle = false;
            List<string> bloques = new();
            string bloqueActual = transaccionPendiente;

            foreach (string lineaRaw in lineas)
            {
                string linea = lineaRaw.Trim();
                string lineaMayus = linea.ToUpper();

                if (lineaMayus.Contains("DETALLE DE OPERACIONES"))
                    enDetalle = true;
                if (enDetalle && DebeTerminarSeccion(lineaMayus))
                    enDetalle = false;

                if (!enDetalle)
                    continue;

                if (lineaMayus.Contains("FECHA") && lineaMayus.Contains("CONCEPTO"))
                    continue;
                if (Regex.IsMatch(lineaMayus, @"PÁGINA:\s*\d+\s*DE\s*\d+"))
                    continue;
                if (lineaMayus.Contains("CIFIBANAMEX"))
                    continue;

                if (Regex.IsMatch(linea, @"^\d{1,2}[\s-]+(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\b", RegexOptions.IgnoreCase))
                {
                    if (!string.IsNullOrEmpty(bloqueActual))
                        bloques.Add(bloqueActual);
                    bloqueActual = linea;
                }
                else
                {
                    bloqueActual += " " + linea;
                }
            }

            ProcesarBloques(bloques, movimientos, anio);
            return bloqueActual;
        }

        private bool DebeTerminarSeccion(string linea)
        {
            return linea.Contains("COMISIONES COBRADAS") ||
                linea.Contains("RESUMEN") ||
                linea.Contains("ESTADO DE CUENTA") ||
                linea.Contains("CLIENTE:") ||
                linea.Contains("SUBTOTALES") ||
                linea.Contains("SALDO MINIMO REQUERIDO");
        }

        private void ProcesarBloques(List<string> bloques, List<MovimientoBanamex> movimientos, int anio)
        {
            foreach (string bloque in bloques)
            {
                // Se extrae la fecha y el resto del bloque; se asume que comienza con "dd MES"
                var match = Regex.Match(bloque, @"^(?<dia>\d{1,2})\s+(?<mes>[A-Z]{3})\s+(?<resto>.*)$", RegexOptions.IgnoreCase);
                if (!match.Success) continue;

                string dia = match.Groups["dia"].Value;
                string mes = match.Groups["mes"].Value;
                string resto = match.Groups["resto"].Value;

                string deposit = "";
                string saldo = "";
                string concepto = "";

                // Si el bloque contiene "IMPORTE:" (con signo $)
                if (resto.IndexOf("IMPORTE:", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    int idx = resto.IndexOf("IMPORTE:", StringComparison.OrdinalIgnoreCase);
                    // Todo lo anterior a "IMPORTE:" es parte del concepto
                    string conceptPart = resto.Substring(0, idx).Trim();
                    string amountPart = resto.Substring(idx);
                    // Eliminar "IMPORTE:" y el signo $ de la parte de montos
                    amountPart = Regex.Replace(amountPart, @"IMPORTE:\s*\$?\s*", "", RegexOptions.IgnoreCase);
                    // Se extraen los montos sin el signo $, usando lookbehind negativo
                    var amountsAfter = Regex.Matches(amountPart, @"(?<!\$)\b\d{1,3}(?:,\d{3})*\.\d{2}\b");
                    List<string> newAmounts = amountsAfter.Cast<Match>().Select(m => m.Value).ToList();

                    if (newAmounts.Count == 0)
                    {
                        deposit = "";
                        saldo = "";
                    }
                    else if (newAmounts.Count == 1)
                    {
                        deposit = newAmounts[0];
                        saldo = "";
                    }
                    else
                    {
                        deposit = newAmounts[0];
                        saldo = newAmounts[1];
                    }
                    concepto = conceptPart;
                }
                else
                {
                    // Fallback: extraer todos los montos (ignorando aquellos cercanos a "REF", "RFB" o "DIVISA")
                    var allMontos = Regex.Matches(resto, @"\b\d{1,3}(?:,\d{3})*\.\d{2}\b");
                    List<string> listaMontos = allMontos.Cast<Match>().Select(m => m.Value).ToList();
                    List<string> filteredMontos = new();
                    foreach (string mo in listaMontos)
                    {
                        int index = resto.IndexOf(mo);
                        int start = Math.Max(0, index - 15);
                        int length = Math.Min(30, resto.Length - start);
                        string snippet = resto.Substring(start, length).ToUpper();
                        if (snippet.Contains("REF") || snippet.Contains("RFB") || snippet.Contains("DIVISA"))
                            continue;
                        filteredMontos.Add(mo);
                    }

                    if (filteredMontos.Count == 0)
                    {
                        deposit = "";
                        saldo = "";
                    }
                    else if (filteredMontos.Count == 1)
                    {
                        deposit = filteredMontos[0];
                        saldo = "";
                    }
                    else
                    {
                        deposit = filteredMontos[0];
                        saldo = filteredMontos[1];
                    }
                    // Se eliminan los montos extraídos para limpiar el concepto
                    foreach (string mVal in filteredMontos)
                    {
                        resto = resto.Replace(mVal, "").Trim();
                    }
                    concepto = resto.Trim();
                }

                MovimientoBanamex mov = new MovimientoBanamex
                {
                    FECHA = $"{dia}-{mes}",
                    ANIO = anio,
                    CONCEPTO = concepto,
                    RETIROS_DEPOSITOS = deposit,
                    SALDO = saldo
                };

                movimientos.Add(mov);
            }
        }


        // -----------------------------------------------------------------
        // Endpoint para procesar PDFs tipo Banorte
        // -----------------------------------------------------------------
        [HttpPost("banorte")]
        [ProducesResponseType(typeof(List<MovimientoBanorte>), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public IActionResult ProcesarBanorte(
            [FromForm][Required] IFormFile file,
            [FromForm][Required] int anio)
        {
            if (file.Length == 0)
                return BadRequest("No se proporcionó un archivo PDF válido.");

            try
            {
                var text = ExtraerTextoPdf(file);
                var movimientos = ExtraerMovimientosBanorte(text, anio);
                return Ok(movimientos);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al procesar PDF Banorte: {ex.Message}");
            }
        }

        private string ExtraerTextoPdf(IFormFile file)
        {
            using var stream = file.OpenReadStream();
            using var reader = new PdfReader(stream);
            using var pdfDoc = new PdfDocument(reader);

            var fullText = new StringBuilder();
            for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
            {
                var strategy = new SimpleTextExtractionStrategy();
                string pageText = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(i), strategy);
                fullText.Append(ProcesarTexto(pageText));
            }
            return fullText.ToString();
        }

        private string ProcesarTexto(string text)
        {
            var processedText = new StringBuilder();
            var lines = text.Split('\n');
            foreach (var line in lines)
            {
                string cleanedLine = Regex.Replace(line.Trim(), @"\s{2,}", " ");
                if (Regex.IsMatch(cleanedLine, @"^\d{2}-[A-Z]{3}-\d{2}"))
                    processedText.AppendLine("\n" + cleanedLine);
                else
                    processedText.AppendLine(cleanedLine);
            }
            return processedText.ToString();
        }

        private List<MovimientoBanorte> ExtraerMovimientosBanorte(string text, int anio)
        {
            var movimientos = new List<MovimientoBanorte>();
            var ignorePatterns = new[]
            {
                "Línea Directa para su empresa:",
                "Visita nuestra página:",
                "Banco Mercantil del Norte S.A.",
                "Institución de Banca Múltiple Grupo Financiero Banorte",
                "ESTADO DE CUENTA / ENLACE NEGOCIOS PFAE"
            };

            string[] bloques = text.Split(new[] { "DETALLE DE MOVIMIENTOS (PESOS)" }, StringSplitOptions.None);
            
            foreach (var bloque in bloques.Skip(1))
            {
                int indexFin = bloque.IndexOf("OTROS");
                string contenido = indexFin != -1 ? bloque.Substring(0, indexFin) : bloque;

                var lineas = contenido.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                ProcesarLineas(lineas, movimientos, anio, ignorePatterns);
            }

            return movimientos;
        }

        private void ProcesarLineas(string[] lineas, List<MovimientoBanorte> movimientos, int anio, string[] ignorePatterns)
        {
            MovimientoBanorte current = null;
            decimal saldoAnterior = 0;

            foreach (var linea in lineas)
            {
                if (ignorePatterns.Any(p => linea.Contains(p))) continue;

                var matchFecha = Regex.Match(linea, @"^(\d{2}-[A-Z]{3}-\d{2})(.*)");
                if (matchFecha.Success)
                {
                    if (current != null)
                    {
                        FinalizarMovimiento(current, saldoAnterior);
                        movimientos.Add(current);
                        saldoAnterior = decimal.Parse(current.Saldo, NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture);
                    }

                    current = new MovimientoBanorte
                    {
                        Fecha = matchFecha.Groups[1].Value.Trim(),
                        Descripcion = matchFecha.Groups[2].Value.Trim(),
                        Anio = anio
                    };

                    ProcesarMontos(current);
                }
                else if (current != null)
                {
                    current.Descripcion += " " + linea.Trim();
                    ProcesarMontos(current);
                }
            }

            if (current != null)
            {
                FinalizarMovimiento(current, saldoAnterior);
                movimientos.Add(current);
            }
        }

        private void ProcesarMontos(MovimientoBanorte movimiento)
        {
            var partes = movimiento.Descripcion.Split(' ');
            movimiento.Descripcion = "";
            var montos = new List<string>();

            foreach (var parte in partes)
            {
                if (Regex.IsMatch(parte, @"^\d{1,3}(,\d{3})*\.\d{2}$"))
                    montos.Add(parte);
                else
                    movimiento.Descripcion += parte + " ";
            }

            movimiento.Descripcion = movimiento.Descripcion.Trim();

            if (montos.Count > 0)
            {
                movimiento.Saldo = montos.Last();
                if (montos.Count == 2)
                {
                    movimiento.MontoDeposito = montos[0];
                    movimiento.MontoRetiro = "";
                }
                else if (montos.Count == 1)
                {
                    movimiento.MontoDeposito = "";
                    movimiento.MontoRetiro = "";
                }
            }
        }

        private void FinalizarMovimiento(MovimientoBanorte movimiento, decimal saldoAnterior)
        {
            if (!string.IsNullOrEmpty(movimiento.Saldo))
            {
                var saldoActual = decimal.Parse(movimiento.Saldo, NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture);
                
                if (!string.IsNullOrEmpty(movimiento.MontoDeposito))
                {
                    var monto = decimal.Parse(movimiento.MontoDeposito, NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture);
                    if (saldoActual < saldoAnterior)
                    {
                        movimiento.MontoRetiro = movimiento.MontoDeposito;
                        movimiento.MontoDeposito = "";
                    }
                }
            }
        }
    
        // -----------------------------------------------------------------
        // NUEVO: Endpoint general para exportar a Excel para todos los bancos.
        // -----------------------------------------------------------------
        [HttpPost("excel")]
        [ProducesResponseType(typeof(FileContentResult), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public IActionResult ExportToExcel(
            [FromForm][Required] IFormFile file,
            [FromForm][Required] int anio,
            [FromForm][Required] string banco, // Valores: "bbva", "banbajio", "banamex" o "banorte"
            [FromForm] bool confirmacion = false)
        {
            if (file.Length == 0)
                return BadRequest("No se proporcionó un archivo PDF válido.");

            try
            {
                if (banco.Equals("bbva", StringComparison.OrdinalIgnoreCase))
                {
                    List<MovimientoBBVA> movimientos = new();
                    using (var memoryStream = new MemoryStream())
                    {
                        file.CopyTo(memoryStream);
                        memoryStream.Position = 0;
                        using (var reader = new PdfReader(memoryStream))
                        using (var pdfDoc = new PdfDocument(reader))
                        {
                            for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
                            {
                                var strategy = new SimpleTextExtractionStrategy();
                                string pageText = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy);
                                movimientos.AddRange(ExtraerMovimientosBBVA(pageText, anio));
                            }
                        }
                    }
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Movimientos BBVA");
                        worksheet.Cell(1, 1).Value = "OPER";
                        worksheet.Cell(1, 2).Value = "LIQ";
                        worksheet.Cell(1, 3).Value = "ANIO";
                        worksheet.Cell(1, 4).Value = "COD_DESCRIPCION";
                        worksheet.Cell(1, 5).Value = "REFERENCIA";
                        worksheet.Cell(1, 6).Value = "CARGOS_ABONOS";
                        worksheet.Cell(1, 7).Value = "OPERACION";
                        worksheet.Cell(1, 8).Value = "LIQUIDACION";

                        int row = 2;
                        foreach (var mov in movimientos)
                        {
                            worksheet.Cell(row, 1).Value = mov.OPER;
                            worksheet.Cell(row, 2).Value = mov.LIQ;
                            worksheet.Cell(row, 3).Value = mov.ANIO;
                            worksheet.Cell(row, 4).Value = mov.COD_DESCRIPCION;
                            worksheet.Cell(row, 5).Value = mov.REFERENCIA;
                            worksheet.Cell(row, 6).Value = mov.CARGOS_ABONOS;
                            worksheet.Cell(row, 7).Value = mov.OPERACION;
                            worksheet.Cell(row, 8).Value = mov.LIQUIDACION;
                            row++;
                        }

                        using (var stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);
                            var content = stream.ToArray();
                            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "MovimientosBBVA.xlsx");
                        }
                    }
                }
                else if (banco.Equals("banbajio", StringComparison.OrdinalIgnoreCase))
                {
                    List<MovimientoBanBajio> movimientos = new();
                    using (var memoryStream = new MemoryStream())
                    {
                        file.CopyTo(memoryStream);
                        memoryStream.Position = 0;
                        using (var reader = new PdfReader(memoryStream))
                        using (var pdfDoc = new PdfDocument(reader))
                        {
                            for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
                            {
                                var strategy = new SimpleTextExtractionStrategy();
                                string pageText = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy);
                                movimientos.AddRange(ExtraerMovimientosBanBajio(pageText, anio));
                            }
                        }
                    }
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Movimientos BanBajio");
                        worksheet.Cell(1, 1).Value = "FECHA";
                        worksheet.Cell(1, 2).Value = "AÑO";
                        worksheet.Cell(1, 3).Value = "REF_DOCT";
                        worksheet.Cell(1, 4).Value = "DESCRIPCION";
                        worksheet.Cell(1, 5).Value = "DEPOSITOS/RETIROS";
                        worksheet.Cell(1, 6).Value = "SALDO";

                        int row = 2;
                        foreach (var mov in movimientos)
                        {
                            worksheet.Cell(row, 1).Value = mov.FECHA;
                            worksheet.Cell(row, 2).Value = mov.ANIO;
                            worksheet.Cell(row, 3).Value = mov.REF_DOCT;
                            worksheet.Cell(row, 4).Value = mov.DESCRIPCION;
                            worksheet.Cell(row, 5).Value = mov.DEPOSITOS_RETIROS;
                            worksheet.Cell(row, 6).Value = mov.SALDO;
                            row++;
                        }

                        using (var stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);
                            var content = stream.ToArray();
                            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "MovimientosBanBajio.xlsx");
                        }
                    }
                }
                else if (banco.Equals("banamex", StringComparison.OrdinalIgnoreCase))
                {
                    List<MovimientoBanamex> movimientos = new();
                    string transaccionPendiente = string.Empty;
                    using (var memoryStream = new MemoryStream())
                    {
                        file.CopyTo(memoryStream);
                        memoryStream.Position = 0;
                        using (var reader = new PdfReader(memoryStream))
                        using (var pdfDoc = new PdfDocument(reader))
                        {
                            for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
                            {
                                var strategy = new SimpleTextExtractionStrategy();
                                string pageText = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy);
                                pageText = ProcessTextBanamex(pageText);
                                transaccionPendiente = ExtractMovementsFromPageBanamex(pageText, movimientos, anio, transaccionPendiente);
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(transaccionPendiente))
                    {
                        var lines = transaccionPendiente.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        ProcesarBloques(lines, movimientos, anio);
                    }

                    if (movimientos.Count > 0 && !confirmacion)
                    {
                        var last = movimientos.Last();
                        return BadRequest(new
                        {
                            message = "Verifique el último registro antes de exportar. Envíe confirmacion=true para proceder.",
                            lastRecord = new
                            {
                                last.FECHA,
                                last.ANIO,
                                last.CONCEPTO,
                                last.RETIROS_DEPOSITOS,
                                last.SALDO
                            }
                        });
                    }

                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Movimientos Banamex");
                        worksheet.Cell(1, 1).Value = "FECHA";
                        worksheet.Cell(1, 2).Value = "AÑO";
                        worksheet.Cell(1, 3).Value = "CONCEPTO";
                        worksheet.Cell(1, 4).Value = "RETIROS/DEPOSITOS";
                        worksheet.Cell(1, 5).Value = "SALDO";

                        int row = 2;
                        foreach (var mov in movimientos)
                        {
                            worksheet.Cell(row, 1).Value = mov.FECHA;
                            worksheet.Cell(row, 2).Value = mov.ANIO;
                            worksheet.Cell(row, 3).Value = mov.CONCEPTO;
                            worksheet.Cell(row, 4).Value = mov.RETIROS_DEPOSITOS;
                            worksheet.Cell(row, 5).Value = mov.SALDO;
                            row++;
                        }

                        using (var stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);
                            var content = stream.ToArray();
                            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "MovimientosBanamex.xlsx");
                        }
                    }
                }
                else if (banco.Equals("banorte", StringComparison.OrdinalIgnoreCase))
                {
                    List<MovimientoBanorte> movimientos = new();
                    using (var memoryStream = new MemoryStream())
                    {
                        file.CopyTo(memoryStream);
                        memoryStream.Position = 0;
                        using (var reader = new PdfReader(memoryStream))
                        using (var pdfDoc = new PdfDocument(reader))
                        {
                            for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
                            {
                                var strategy = new SimpleTextExtractionStrategy();
                                string pageText = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy);
                                movimientos.AddRange(ExtraerMovimientosBanorte(pageText, anio));
                            }
                        }
                    }
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Movimientos Banorte");
                        worksheet.Cell(1, 1).Value = "FECHA";
                        worksheet.Cell(1, 2).Value = "DESCRIPCION";
                        worksheet.Cell(1, 3).Value = "MONTO DEPOSITO";
                        worksheet.Cell(1, 4).Value = "MONTO RETIRO";
                        worksheet.Cell(1, 5).Value = "SALDO";
                        worksheet.Cell(1, 6).Value = "AÑO";

                        int row = 2;
                        foreach (var mov in movimientos)
                        {
                            worksheet.Cell(row, 1).Value = mov.Fecha;
                            worksheet.Cell(row, 2).Value = mov.Descripcion;
                            worksheet.Cell(row, 3).Value = mov.MontoDeposito;
                            worksheet.Cell(row, 4).Value = mov.MontoRetiro;
                            worksheet.Cell(row, 5).Value = mov.Saldo; // Si aplica, o agrega si lo necesitas
                            worksheet.Cell(row, 6).Value = mov.Anio;
                            row++;
                         }

                        using (var stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);
                            var content = stream.ToArray();
                            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "MovimientosBanorte.xlsx");
                        }
                    }
                }
                else
                {
                    return BadRequest("Banco no soportado. Use 'bbva', 'banbajio', 'banamex' o 'banorte'.");
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al exportar a Excel: {ex.Message}");
            }
        }
        // *****************************************************************
        // Modelos
        // *****************************************************************
        public class MovimientoBBVA
        {
             [Required]
            public string? OPER { get; set; }
            [Required]
            public string? LIQ { get; set; }
            [Required]
            public int ANIO { get; set; }
            public string? COD_DESCRIPCION { get; set; }
            public string? REFERENCIA { get; set; }
            public string? CARGOS_ABONOS { get; set; }
            public string? OPERACION { get; set; }
            public string? LIQUIDACION { get; set; }
        }

        public class MovimientoBanBajio
        {
            [Required]
            public string? FECHA { get; set; }
            [Required]
            public int ANIO { get; set; }
            public string? REF_DOCT { get; set; }
            public string? DESCRIPCION { get; set; }
            public string? DEPOSITOS_RETIROS { get; set; }
            public string? SALDO { get; set; }
        }

        public class MovimientoBanamex
        {
            [Required]
            public string? FECHA { get; set; }
            [Required]
            public int ANIO { get; set; }
            public string? CONCEPTO { get; set; }
            public string? RETIROS_DEPOSITOS { get; set; }
            public string? SALDO { get; set; }
        }

        public class MovimientoBanorte
        {
             [Required]
            public string Fecha { get; set; }
            [Required]
            public string Descripcion { get; set; }
            public string MontoDeposito { get; set; }
            public string MontoRetiro { get; set; }
            public string Saldo { get; set; }
            [Required]
            public int Anio { get; set; 
        }
    }
    }
}
