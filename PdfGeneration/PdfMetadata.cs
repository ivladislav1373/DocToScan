using System;

namespace DocToScan.PdfGeneration;

/// <summary>
/// Представляет метаданные PDF документа.
/// </summary>
public class PdfMetadata
{
    /// <summary>
    /// Получает или задает название документа.
    /// </summary>
    public string Title { get; set; }

    /// <summary>
    /// Получает или задает автора документа.
    /// </summary>
    public string Author { get; set; }

    /// <summary>
    /// Получает или задает тему документа.
    /// </summary>
    public string Subject { get; set; }

    /// <summary>
    /// Получает или задает ключевые слова.
    /// </summary>
    public string Keywords { get; set; }

    /// <summary>
    /// Получает или задает название программы, создавшей документ.
    /// </summary>
    public string Creator { get; set; }

    /// <summary>
    /// Получает или задает название программы, использованной для создания документа.
    /// </summary>
    public string Producer { get; set; }

    /// <summary>
    /// Получает или задает дату создания документа.
    /// </summary>
    public DateTime? CreationDate { get; set; }

    /// <summary>
    /// Получает или задает дату последнего изменения.
    /// </summary>
    public DateTime? ModificationDate { get; set; }

    /// <summary>
    /// Создает метаданные по умолчанию для DocToScan.
    /// </summary>
    public static PdfMetadata CreateDefault()
    {
        return new PdfMetadata
        {
            Title = "Документ",
            Author = "Vlad",
            Creator = "Scanner v1.0",
            Producer = "PDFsharp",
            CreationDate = DateTime.Now,
            ModificationDate = DateTime.Now,
            Keywords = "scan, pdf, document",
            Subject = "Отсканированный документ"
        };
    }
}