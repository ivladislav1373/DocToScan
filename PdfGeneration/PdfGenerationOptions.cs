using DocToScan.Configuration;

namespace DocToScan.PdfGeneration;

/// <summary>
/// Представляет опции для настройки процесса генерации PDF документа.
/// </summary>
public class PdfGenerationOptions
{
    /// <summary>
    /// Качество JPEG сжатия (1-100)
    /// </summary>
    public int JpegQuality { get; set; } = 85;

    /// <summary>
    /// Название документа
    /// </summary>
    public string DocumentTitle { get; set; }

    /// <summary>
    /// Автор документа
    /// </summary>
    public string Author { get; set; }

    /// <summary>
    /// Создатель документа
    /// </summary>
    public string Creator { get; set; }

    /// <summary>
    /// Ключевые слова
    /// </summary>
    public string Keywords { get; set; }

    /// <summary>
    /// Тема документа
    /// </summary>
    public string Subject { get; set; }

    /// <summary>
    /// Добавлять дату создания
    /// </summary>
    public bool AddCreationDate { get; set; } = true;

    /// <summary>
    /// Сжимать содержимое
    /// </summary>
    public bool CompressContent { get; set; } = true;

    /// <summary>
    /// Создаёт опции на основе конфигурации
    /// </summary>
    public static PdfGenerationOptions FromConfig(Config config, string documentTitle = null)
    {
        return new PdfGenerationOptions
        {
            JpegQuality = config.ImageQuality.JpegCompression,
            DocumentTitle = documentTitle ?? "Скан-копия документа",
            Author = config.PdfMetadata.Author,
            Creator = config.PdfMetadata.Creator,
            Keywords = config.PdfMetadata.Keywords,
            Subject = config.PdfMetadata.Subject,
            AddCreationDate = config.PdfMetadata.AddCreationDate,
            CompressContent = config.PdfMetadata.CompressContent
        };
    }
}