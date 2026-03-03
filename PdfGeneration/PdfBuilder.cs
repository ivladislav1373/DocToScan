using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using DocToScan.Logging;
using PdfSharp.Pdf;
using PdfSharp.Drawing;

namespace DocToScan.PdfGeneration;

/// <summary>
/// Построитель PDF документов из изображений.
/// Отвечает за создание многостраничного PDF файла из коллекции изображений.
/// </summary>
/// <remarks>
/// Использует библиотеку PDFsharp для создания PDF документа.
/// Поддерживает настройку качества сжатия JPEG и оптимизацию размера файла.
/// </remarks>
public class PdfBuilder : IDisposable
{
    private readonly ILogger _logger;
    private readonly PdfGenerationOptions _options;
    private readonly PdfDocument _document;
    private readonly List<IDisposable> _resources;
    private bool _disposed;

    /// <summary>
    /// Инициализирует новый экземпляр класса PdfBuilder.
    /// </summary>
    /// <param name="logger">Логгер для записи сообщений о процессе создания PDF.</param>
    /// <param name="options">Опции генерации PDF документа.</param>
    /// <exception cref="ArgumentNullException">
    /// Выбрасывается, если logger или options равны null.
    /// </exception>
    public PdfBuilder(ILogger logger, PdfGenerationOptions options = null)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _options = options ?? new PdfGenerationOptions();
        _document = new PdfDocument();
        _resources = new List<IDisposable>();

        _logger.Debug("PDF builder инициализирован");
    }

    /// <summary>
    /// Получает количество страниц в текущем PDF документе.
    /// </summary>
    public int PageCount => _document.PageCount;

    /// <summary>
    /// Добавляет новую страницу в PDF документ из изображения.
    /// </summary>
    /// <param name="image">Изображение для добавления в качестве страницы.</param>
    /// <exception cref="ArgumentNullException">Выбрасывается, если image равен null.</exception>
    /// <exception cref="InvalidOperationException">Выбрасывается при ошибке добавления страницы.</exception>
    /// <remarks>
    /// Размер страницы автоматически подстраивается под размер изображения с учетом DPI.
    /// Изображение сжимается с качеством, указанным в опциях.
    /// </remarks>
    public void AddPage(Image image)
    {
        if (image == null)
            throw new ArgumentNullException(nameof(image));

        try
        {
            _logger.Debug($"Добавление страницы {_document.PageCount + 1}, размер изображения: {image.Width}x{image.Height}");

            // Создаем новую страницу с размером, соответствующим изображению
            var page = _document.AddPage();

            // Конвертируем размер из пикселей в пункты (1/72 дюйма)
            // Используем горизонтальное разрешение изображения для расчета
            page.Width = XUnit.FromPoint(image.Width * 72 / image.HorizontalResolution);
            page.Height = XUnit.FromPoint(image.Height * 72 / image.VerticalResolution);

            _logger.Debug($"  Размер страницы в пунктах: {page.Width:F2}x{page.Height:F2}");

            using (var gfx = XGraphics.FromPdfPage(page))
            {
                // Сохраняем изображение во временный поток с нужным качеством
                var imageStream = CompressImage(image);

                // Загружаем в XImage и рисуем на странице
                var xImage = XImage.FromStream(imageStream);
                _resources.Add(xImage); // Для последующего освобождения

                // Рисуем изображение на всю страницу
                gfx.DrawImage(xImage, 0, 0, page.Width.Point, page.Height.Point);
            }

            _logger.Debug($"Страница {_document.PageCount} успешно добавлена");
        }
        catch (Exception ex)
        {
            _logger.Error($"Ошибка при добавлении страницы: {ex.Message}");
            throw new InvalidOperationException("Не удалось добавить страницу в PDF", ex);
        }
    }

    /// <summary>
    /// Сжимает изображение с использованием JPEG кодека.
    /// </summary>
    /// <param name="image">Исходное изображение.</param>
    /// <returns>Поток памяти со сжатым изображением.</returns>
    private MemoryStream CompressImage(Image image)
    {
        var encoderParams = new EncoderParameters(1);
        encoderParams.Param[0] = new EncoderParameter(
            Encoder.Quality,
            _options.JpegQuality);

        var codecInfo = GetEncoderInfo("image/jpeg");

        var ms = new MemoryStream();
        image.Save(ms, codecInfo, encoderParams);
        ms.Position = 0;

        _resources.Add(ms);

        _logger.Debug($"  Изображение сжато, качество JPEG: {_options.JpegQuality}");
        return ms;
    }

    /// <summary>
    /// Получает кодировщик для указанного MIME-типа.
    /// </summary>
    /// <param name="mimeType">MIME-тип (например, "image/jpeg").</param>
    /// <returns>Кодировщик изображений.</returns>
    /// <exception cref="Exception">Выбрасывается, если кодировщик не найден.</exception>
    private ImageCodecInfo GetEncoderInfo(string mimeType)
    {
        var encoders = ImageCodecInfo.GetImageEncoders();
        foreach (var encoder in encoders)
        {
            if (encoder.MimeType == mimeType)
                return encoder;
        }

        throw new Exception($"Не найден кодировщик для {mimeType}");
    }

    /// <summary>
    /// Сохраняет PDF документ в файл.
    /// </summary>
    /// <param name="outputPath">Путь для сохранения PDF файла.</param>
    /// <exception cref="ArgumentException">Выбрасывается, если путь пустой или некорректный.</exception>
    /// <exception cref="InvalidOperationException">Выбрасывается при ошибке сохранения.</exception>
    public void Save(string outputPath)
    {
        if (string.IsNullOrWhiteSpace(outputPath))
            throw new ArgumentException("Путь к выходному файлу не может быть пустым", nameof(outputPath));

        try
        {
            _logger.Debug($"Сохранение PDF в файл: {outputPath}");
            _logger.Info($"    Страниц в документе: {_document.PageCount}");

            // Настройка метаданных документа
            _document.Info.Title = _options.DocumentTitle ?? "Скан-копия документа";
            _document.Info.Author = _options.Author ?? "DocToScan";
            _document.Info.Creator = "DocToScan v1.0";
            _document.Info.Keywords = _options.Keywords ?? "scan, pdf, document";

            _document.Save(outputPath);

            _logger.Success($"    PDF сохранен, размер: {GetFileSize(outputPath)}");
        }
        catch (Exception ex)
        {
            _logger.Error($"Ошибка при сохранении PDF: {ex.Message}");
            throw new InvalidOperationException($"Не удалось сохранить PDF файл: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Сохраняет PDF документ в поток.
    /// </summary>
    /// <param name="stream">Поток для сохранения.</param>
    /// <exception cref="ArgumentNullException">Выбрасывается, если stream равен null.</exception>
    public void Save(Stream stream)
    {
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));

        try
        {
            _document.Save(stream);
            _logger.Debug("PDF сохранен в поток");
        }
        catch (Exception ex)
        {
            _logger.Error($"Ошибка при сохранении PDF в поток: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Получает форматированный размер файла.
    /// </summary>
    private string GetFileSize(string filePath)
    {
        var fi = new FileInfo(filePath);
        if (fi.Length < 1024)
            return $"{fi.Length} B";
        if (fi.Length < 1024 * 1024)
            return $"{fi.Length / 1024.0:F1} KB";
        return $"{fi.Length / (1024.0 * 1024.0):F1} MB";
    }

    /// <summary>
    /// Освобождает ресурсы, используемые построителем PDF.
    /// </summary>
    public void Dispose()
    {
        if (!_disposed)
        {
            _logger.Debug("Освобождение ресурсов PDF builder");

            // Освобождаем все зарегистрированные ресурсы
            foreach (var resource in _resources)
            {
                try
                {
                    resource?.Dispose();
                }
                catch (Exception ex)
                {
                    _logger.Debug($"Ошибка при освобождении ресурса: {ex.Message}");
                }
            }
            _resources.Clear();

            // Освобождаем документ
            _document?.Dispose();

            _disposed = true;
            _logger.Debug("Ресурсы PDF builder освобождены");
        }
    }
}