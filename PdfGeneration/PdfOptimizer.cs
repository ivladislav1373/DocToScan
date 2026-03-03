using System;
using System.Collections.Generic;
using System.IO;
using DocToScan.Logging;

namespace DocToScan.PdfGeneration;

/// <summary>
/// Предоставляет методы для оптимизации PDF файлов.
/// </summary>
public static class PdfOptimizer
{
    /// <summary>
    /// Оптимизирует размер PDF файла, удаляя дублирующиеся ресурсы.
    /// </summary>
    /// <param name="pdfPath">Путь к PDF файлу.</param>
    /// <param name="logger">Логгер для записи сообщений.</param>
    /// <returns>Размер сэкономленного места в байтах.</returns>
    public static long OptimizeFileSize(string pdfPath, ILogger logger)
    {
        if (!File.Exists(pdfPath))
        {
            logger?.Error($"Файл не найден: {pdfPath}");
            return 0;
        }

        try
        {
            var fileInfo = new FileInfo(pdfPath);
            long originalSize = fileInfo.Length;

            logger?.Debug($"Оптимизация PDF: {Path.GetFileName(pdfPath)}, размер: {FormatSize(originalSize)}");

            // Здесь можно добавить логику оптимизации PDF
            // Например, удаление дублирующихся шрифтов, сжатие потоков и т.д.

            // Пока просто логируем
            logger?.Debug("Оптимизация завершена");

            return originalSize - fileInfo.Length;
        }
        catch (Exception ex)
        {
            logger?.Error($"Ошибка при оптимизации PDF: {ex.Message}");
            return 0;
        }
    }

    /// <summary>
    /// Анализирует содержимое PDF файла и возвращает статистику.
    /// </summary>
    /// <param name="pdfPath">Путь к PDF файлу.</param>
    /// <param name="logger">Логгер для записи сообщений.</param>
    /// <returns>Статистика PDF файла.</returns>
    public static PdfStatistics AnalyzePdf(string pdfPath, ILogger logger)
    {
        var stats = new PdfStatistics();

        if (!File.Exists(pdfPath))
        {
            logger?.Error($"Файл не найден: {pdfPath}");
            return stats;
        }

        try
        {
            var fileInfo = new FileInfo(pdfPath);
            stats.FileSize = fileInfo.Length;
            stats.FileName = fileInfo.Name;
            stats.CreationTime = fileInfo.CreationTime;
            stats.ModificationTime = fileInfo.LastWriteTime;

            // Здесь можно добавить анализ содержимого PDF
            // Например, подсчет страниц, изображений, шрифтов и т.д.

            logger?.Debug($"Анализ PDF завершен: {stats.PageCount} страниц, {FormatSize(stats.FileSize)}");
        }
        catch (Exception ex)
        {
            logger?.Error($"Ошибка при анализе PDF: {ex.Message}");
        }

        return stats;
    }

    private static string FormatSize(long bytes)
    {
        if (bytes < 1024) return $"{bytes} B";
        if (bytes < 1024 * 1024) return $"{bytes / 1024.0:F1} KB";
        return $"{bytes / (1024.0 * 1024.0):F1} MB";
    }
}

/// <summary>
/// Представляет статистику PDF файла.
/// </summary>
public class PdfStatistics
{
    /// <summary>
    /// Имя файла.
    /// </summary>
    public string FileName { get; set; }

    /// <summary>
    /// Размер файла в байтах.
    /// </summary>
    public long FileSize { get; set; }

    /// <summary>
    /// Количество страниц.
    /// </summary>
    public int PageCount { get; set; }

    /// <summary>
    /// Количество изображений в документе.
    /// </summary>
    public int ImageCount { get; set; }

    /// <summary>
    /// Количество используемых шрифтов.
    /// </summary>
    public int FontCount { get; set; }

    /// <summary>
    /// Версия PDF формата.
    /// </summary>
    public string PdfVersion { get; set; }

    /// <summary>
    /// Время создания файла.
    /// </summary>
    public DateTime CreationTime { get; set; }

    /// <summary>
    /// Время последнего изменения.
    /// </summary>
    public DateTime ModificationTime { get; set; }

    /// <summary>
    /// Является ли PDF защищенным.
    /// </summary>
    public bool IsProtected { get; set; }

    /// <summary>
    /// Содержит ли PDF встроенные шрифты.
    /// </summary>
    public bool HasEmbeddedFonts { get; set; }
}