using System;

namespace DocToScan.PdfGeneration;

/// <summary>
/// Представляет опции для настройки процесса генерации PDF документа.
/// </summary>
public class PdfGenerationOptions
{
    /// <summary>
    /// Получает или задает качество JPEG сжатия для изображений.
    /// Значение от 1 (низкое качество, малый размер) до 100 (максимальное качество, большой размер).
    /// </summary>
    public int JpegQuality { get; set; } = 85;

    /// <summary>
    /// Получает или задает название документа, которое будет сохранено в метаданных PDF.
    /// </summary>
    public string DocumentTitle { get; set; }

    /// <summary>
    /// Получает или задает автора документа в метаданных PDF.
    /// </summary>
    public string Author { get; set; } = "DocToScan";

    /// <summary>
    /// Получает или задает ключевые слова для поиска документа.
    /// </summary>
    public string Keywords { get; set; } = "scan, pdf";

    /// <summary>
    /// Получает или задает значение, указывающее, нужно ли сжимать текстовые объекты в PDF.
    /// </summary>
    public bool CompressContent { get; set; } = true;

    /// <summary>
    /// Получает или задает значение, указывающее, нужно ли добавлять метку времени создания.
    /// </summary>
    public bool AddCreationDate { get; set; } = true;

    /// <summary>
    /// Получает или задает уровень защиты PDF документа.
    /// </summary>
    public SecurityLevel SecurityLevel { get; set; } = SecurityLevel.None;

    /// <summary>
    /// Получает или задает пароль для открытия документа (если требуется защита).
    /// </summary>
    public string UserPassword { get; set; }

    /// <summary>
    /// Получает или задает пароль владельца для изменения прав доступа.
    /// </summary>
    public string OwnerPassword { get; set; }

    /// <summary>
    /// Получает или задает разрешенные действия для защищенного документа.
    /// </summary>
    public PermittedActions PermittedActions { get; set; } = PermittedActions.All;
}

/// <summary>
/// Определяет уровень безопасности PDF документа.
/// </summary>
public enum SecurityLevel
{
    /// <summary>
    /// Без защиты.
    /// </summary>
    None,

    /// <summary>
    /// Защита 40-битным RC4 шифрованием (совместимость с PDF 1.3).
    /// </summary>
    Low,

    /// <summary>
    /// Защита 128-битным AES шифрованием (рекомендуется).
    /// </summary>
    High
}

/// <summary>
/// Определяет разрешенные действия для защищенного PDF документа.
/// </summary>
[Flags]
public enum PermittedActions
{
    /// <summary>
    /// Нет разрешений.
    /// </summary>
    None = 0,

    /// <summary>
    /// Разрешена печать документа.
    /// </summary>
    Print = 1,

    /// <summary>
    /// Разрешено модифицировать содержимое.
    /// </summary>
    ModifyContent = 2,

    /// <summary>
    /// Разрешено копировать текст и графику.
    /// </summary>
    CopyContent = 4,

    /// <summary>
    /// Разрешено добавлять или изменять аннотации.
    /// </summary>
    ModifyAnnotations = 8,

    /// <summary>
    /// Все разрешения.
    /// </summary>
    All = Print | ModifyContent | CopyContent | ModifyAnnotations
}