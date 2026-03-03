using System;
using System.IO; // Добавить эту строку

namespace DocToScan.Processing;

/// <summary>
/// Представляет результат обработки документа.
/// </summary>
public class ProcessingResult
{
    /// <summary>
    /// Получает или задает значение, указывающее, успешно ли выполнена обработка.
    /// </summary>
    public bool Success { get; set; }

    /// <summary>
    /// Получает или задает путь к выходному файлу (скан-версия).
    /// </summary>
    public string OutputPath { get; set; }

    /// <summary>
    /// Получает или задает путь к промежуточному PDF файлу.
    /// Для DOCX это будет сохраненный PDF, для PDF - null.
    /// </summary>
    public string IntermediatePdfPath { get; set; }

    /// <summary>
    /// Получает или задает количество обработанных страниц.
    /// </summary>
    public int PageCount { get; set; }

    /// <summary>
    /// Получает или задает сообщение об ошибке, если обработка не удалась.
    /// </summary>
    public string ErrorMessage { get; set; }

    /// <summary>
    /// Получает информацию о том, был ли создан промежуточный PDF.
    /// </summary>
    public bool HasIntermediatePdf => !string.IsNullOrEmpty(IntermediatePdfPath) && File.Exists(IntermediatePdfPath);
}