using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using DocToScan.Configuration;
using DocToScan.Converters;
using DocToScan.FileSystem;
using DocToScan.ImageProcessing;
using DocToScan.Logging;
using DocToScan.PdfGeneration;

namespace DocToScan.Processing;

/// <summary>
/// Главный класс для обработки документов.
/// Координирует весь процесс конвертации документа в скан-копию.
/// </summary>
public class DocumentProcessor : IDisposable
{
    private readonly Config _config;
    private readonly ILogger _logger;
    private readonly TempFileManager _tempFileManager;
    private bool _disposed;

    /// <summary>
    /// Инициализирует новый экземпляр класса DocumentProcessor.
    /// </summary>
    public DocumentProcessor(Config config, ILogger logger)
    {
        _config = config ?? throw new ArgumentNullException(nameof(config));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _tempFileManager = new TempFileManager(logger);

        if (_config.PageSize.Enable)
        {
            _logger.Info($"  • Размер страницы: {_config.PageSize.WidthMm} x {_config.PageSize.HeightMm} мм");
        }
    }

    /// <summary>
    /// Обрабатывает один файл, преобразуя его в скан-копию.
    /// </summary>
    public ProcessingResult ProcessFile(string inputPath)
    {
        var result = new ProcessingResult();
        string pdfSourcePath = null;
        List<Image> pageImages = null;
        List<Image> processedImages = null;

        try
        {
            // Шаг 1: Получаем PDF источник (и промежуточный PDF для DOCX)
            pdfSourcePath = GetPdfSource(inputPath, out string intermediatePdfPath);
            result.IntermediatePdfPath = intermediatePdfPath;

            // Шаг 2: Генерируем имя для выходного файла (скан-версия)
            result.OutputPath = FileNameHelper.GetUniqueOutputPath(inputPath, "скан");

            // Шаг 3: Рендерим PDF в изображения
            pageImages = RenderPdfToImages(pdfSourcePath, out int pageCount);
            result.PageCount = pageCount;

            // Шаг 4: Применяем эффекты к каждой странице
            processedImages = ApplyEffectsToImages(pageImages);

            // Шаг 5: Собираем итоговый PDF
            BuildPdfFromImages(processedImages, result.OutputPath);

            result.Success = true;
            _logger.Success($"  ✓ Готово: {Path.GetFileName(result.OutputPath)}");

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;

            // В случае ошибки удаляем промежуточный PDF, если он был создан
            CleanupIntermediatePdfOnError(result.IntermediatePdfPath);

            return result;
        }
        finally
        {
            // Очищаем временные изображения
            CleanupImages(pageImages);
            CleanupImages(processedImages);
        }
    }

    /// <summary>
    /// Получает PDF источник для обработки.
    /// </summary>
    private string GetPdfSource(string inputPath, out string intermediatePdfPath)
    {
        string extension = Path.GetExtension(inputPath).ToLowerInvariant();
        intermediatePdfPath = null;

        if (extension == ".docx" || extension == ".doc")
        {
            return ConvertWordToPdf(inputPath, out intermediatePdfPath);
        }

        if (extension == ".pdf")
        {
            return inputPath;
        }

        throw new NotSupportedException($"Неподдерживаемый формат файла: {extension}");
    }

    /// <summary>
    /// Конвертирует Word документ в PDF.
    /// </summary>
    private string ConvertWordToPdf(string wordPath, out string pdfPath)
    {
        _logger.Info("  • Конвертация DOCX → PDF...");

        pdfPath = GenerateIntermediatePdfPath(wordPath);

        using (var converter = new WordToPdfConverter(_logger))
        {
            converter.ConvertToPdf(wordPath, pdfPath);
        }

        _logger.Success($"    ✓ Сохранен промежуточный PDF: {Path.GetFileName(pdfPath)}");
        return pdfPath;
    }

    /// <summary>
    /// Генерирует уникальный путь для промежуточного PDF.
    /// </summary>
    private string GenerateIntermediatePdfPath(string wordPath)
    {
        string directory = Path.GetDirectoryName(wordPath);
        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(wordPath);
        string basePath = Path.Combine(directory, fileNameWithoutExt + ".pdf");

        if (!File.Exists(basePath))
        {
            return basePath;
        }

        int counter = 1;
        while (true)
        {
            string newPath = Path.Combine(directory, $"{fileNameWithoutExt}_{counter}.pdf");
            if (!File.Exists(newPath))
            {
                return newPath;
            }
            counter++;
        }
    }

    /// <summary>
    /// Рендерит PDF страницы в изображения.
    /// </summary>
    private List<Image> RenderPdfToImages(string pdfPath, out int pageCount)
    {
        _logger.Info("  • Генерация изображений...");

        using (var renderer = new PdfToImageRenderer(_logger))
        {
            renderer.LoadPdf(pdfPath);
            pageCount = renderer.PageCount;
            _logger.Info($"    {pageCount} страниц");

            return renderer.RenderAllPages(_config.ImageQuality.Dpi);
        }
    }

    /// <summary>
    /// Применяет эффекты к списку изображений.
    /// </summary>
    private List<Image> ApplyEffectsToImages(List<Image> sourceImages)
    {
        _logger.Info("  • Применение эффектов...");

        var pipeline = new EffectPipeline(_config, _logger);
        var processedImages = new List<Image>(sourceImages.Count);

        foreach (var image in sourceImages)
        {
            var processed = pipeline.Process(image);
            processedImages.Add(processed);
            image.Dispose(); // Освобождаем исходное изображение
        }

        return processedImages;
    }

    /// <summary>
    /// Собирает PDF из изображений.
    /// </summary>
    private void BuildPdfFromImages(List<Image> images, string outputPath)
    {
        _logger.Info("  • Сохранение результата...");

        // Создаём опции на основе конфигурации
        var options = PdfGenerationOptions.FromConfig(
            _config,
            Path.GetFileNameWithoutExtension(outputPath)
        );

        using (var builder = new PdfBuilder(_logger, options))
        {
            foreach (var image in images)
            {
                builder.AddPage(image);
                image.Dispose(); // Освобождаем обработанное изображение после добавления
            }

            builder.Save(outputPath);
        }
    }

    /// <summary>
    /// Очищает список изображений.
    /// </summary>
    private void CleanupImages(List<Image> images)
    {
        if (images == null) return;

        foreach (var image in images)
        {
            image?.Dispose();
        }
    }

    /// <summary>
    /// Удаляет промежуточный PDF в случае ошибки.
    /// </summary>
    private void CleanupIntermediatePdfOnError(string intermediatePdfPath)
    {
        if (intermediatePdfPath != null && File.Exists(intermediatePdfPath))
        {
            try
            {
                File.Delete(intermediatePdfPath);
                _logger.Debug($"Промежуточный PDF удален из-за ошибки: {intermediatePdfPath}");
            }
            catch
            {
                // Игнорируем ошибки удаления
            }
        }
    }

    /// <summary>
    /// Освобождает ресурсы, используемые процессором.
    /// </summary>
    public void Dispose()
    {
        if (!_disposed)
        {
            _tempFileManager?.Dispose();
            _disposed = true;
        }
    }
}