using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using DocToScan.Logging;
using PdfiumViewer;

namespace DocToScan.Converters;

/// <summary>
/// Предоставляет функциональность для рендеринга PDF страниц в изображения.
/// </summary>
public class PdfToImageRenderer : IDisposable
{
    private readonly ILogger _logger;
    private PdfDocument _pdfDocument;
    private bool _disposed;

    /// <summary>
    /// Инициализирует новый экземпляр класса PdfToImageRenderer.
    /// </summary>
    /// <param name="logger">Логгер для записи сообщений.</param>
    public PdfToImageRenderer(ILogger logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <summary>
    /// Получает количество страниц в загруженном PDF документе.
    /// </summary>
    public int PageCount => _pdfDocument?.PageCount ?? 0;

    /// <summary>
    /// Загружает PDF документ для рендеринга.
    /// </summary>
    /// <param name="pdfPath">Путь к PDF файлу.</param>
    /// <exception cref="FileNotFoundException">Выбрасывается, если файл не найден.</exception>
    /// <exception cref="InvalidOperationException">Выбрасывается при ошибке загрузки PDF.</exception>
    public void LoadPdf(string pdfPath)
    {
        if (!File.Exists(pdfPath))
        {
            throw new FileNotFoundException($"PDF файл не найден: {pdfPath}");
        }

        try
        {
            _pdfDocument?.Dispose();
            _pdfDocument = PdfDocument.Load(pdfPath);
            _logger.Debug($"PDF загружен: {pdfPath}, страниц: {PageCount}");
        }
        catch (Exception ex)
        {
            _logger.Error($"Ошибка загрузки PDF: {ex.Message}");
            throw new InvalidOperationException($"Не удалось загрузить PDF файл: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Рендерит все страницы PDF в изображения.
    /// </summary>
    /// <param name="dpi">Разрешение в DPI для рендеринга.</param>
    /// <returns>Список изображений страниц.</returns>
    public List<Image> RenderAllPages(int dpi)
    {
        if (_pdfDocument == null)
        {
            throw new InvalidOperationException("PDF документ не загружен. Сначала вызовите LoadPdf().");
        }

        var images = new List<Image>();

        try
        {
            for (int i = 0; i < _pdfDocument.PageCount; i++)
            {
                var image = RenderPage(i, dpi);
                images.Add(image);

                _logger.Debug($"Страница {i + 1} отрендерена");
            }

            _logger.Info($"    Отрендерено {images.Count} страниц");
            return images;
        }
        catch (Exception ex)
        {
            // Очищаем уже созданные изображения в случае ошибки
            foreach (var img in images)
            {
                img?.Dispose();
            }

            _logger.Error($"Ошибка рендеринга: {ex.Message}");
            throw new InvalidOperationException($"Ошибка рендеринга страницы: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Рендерит одну страницу PDF в изображение.
    /// </summary>
    /// <param name="pageIndex">Индекс страницы.</param>
    /// <param name="dpi">Разрешение в DPI.</param>
    /// <returns>Изображение страницы.</returns>
    private Image RenderPage(int pageIndex, int dpi)
    {
        var pageSize = _pdfDocument.PageSizes[pageIndex];

        // Конвертируем размер из точек (1/72 дюйма) в пиксели с учетом DPI
        float scale = dpi / 72f;
        int width = (int)(pageSize.Width * scale);
        int height = (int)(pageSize.Height * scale);

        // Рендерим страницу с правильным DPI
        using var pageImage = _pdfDocument.Render(
            pageIndex,
            width,
            height,
            dpi,
            dpi,
            PdfRenderFlags.CorrectFromDpi);

        // Создаем копию, чтобы изображение не было удалено при освобождении PdfDocument
        return new Bitmap(pageImage);
    }

    /// <summary>
    /// Освобождает ресурсы, используемые рендерером.
    /// </summary>
    public void Dispose()
    {
        if (!_disposed)
        {
            _pdfDocument?.Dispose();
            _pdfDocument = null;
            _disposed = true;

            // Принудительная сборка мусора для освобождения неуправляемых ресурсов Pdfium
            GC.Collect();
            GC.WaitForPendingFinalizers();

            _logger.Debug("Ресурсы рендерера PDF освобождены");
        }
    }
}