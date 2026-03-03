using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
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

        // Настраиваем пути к нативным библиотекам при создании экземпляра
        ConfigureNativePaths();
    }

    /// <summary>
    /// Настраивает пути к нативным библиотекам Pdfium.
    /// </summary>
    private void ConfigureNativePaths()
    {
        try
        {
            string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;

            string[] possiblePaths = new[]
            {
                Path.Combine(currentDirectory, "x64"),
                currentDirectory,
                Path.Combine(currentDirectory, "runtimes", "win-x64", "native"),
                Path.GetFullPath(Path.Combine(currentDirectory, "..")),
            };

            string currentPath = Environment.GetEnvironmentVariable("PATH") ?? "";

            foreach (string path in possiblePaths.Distinct())
            {
                if (Directory.Exists(path))
                {
                    string pdfiumPath = Path.Combine(path, "pdfium.dll");
                    if (File.Exists(pdfiumPath) && !currentPath.Contains(path))
                    {
                        Environment.SetEnvironmentVariable("PATH", currentPath + ";" + path);
                        _logger.Debug($"Добавлен путь к нативным библиотекам: {path}");
                    }
                }
            }

            // Добавляем корень, если там есть pdfium.dll
            string rootPdfium = Path.Combine(currentDirectory, "pdfium.dll");
            if (File.Exists(rootPdfium) && !currentPath.Contains(currentDirectory))
            {
                Environment.SetEnvironmentVariable("PATH", currentPath + ";" + currentDirectory);
                _logger.Debug($"Добавлен корневой путь: {currentDirectory}");
            }
        }
        catch (Exception ex)
        {
            _logger.Debug($"Ошибка при настройке путей к нативным библиотекам: {ex.Message}");
        }
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

            _logger.Debug($"Загрузка PDF: {pdfPath}");
            _pdfDocument = PdfDocument.Load(pdfPath);

            _logger.Debug($"PDF загружен: {pdfPath}, страниц: {PageCount}");
        }
        catch (DllNotFoundException ex)
        {
            _logger.Error($"Ошибка загрузки нативной библиотеки Pdfium: {ex.Message}");
            _logger.Error("Убедитесь, что файл pdfium.dll находится в папке программы");

            CheckPdfiumPresence();

            throw new InvalidOperationException(
                "Не удалось загрузить pdfium.dll. Библиотека необходима для работы с PDF.", ex);
        }
        catch (Exception ex)
        {
            _logger.Error($"Ошибка загрузки PDF: {ex.Message}");
            throw new InvalidOperationException($"Не удалось загрузить PDF файл: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Проверяет наличие pdfium.dll в различных местах.
    /// </summary>
    private void CheckPdfiumPresence()
    {
        string currentDir = AppDomain.CurrentDomain.BaseDirectory;
        string[] locationsToCheck = new[]
        {
            Path.Combine(currentDir, "pdfium.dll"),
            Path.Combine(currentDir, "x64", "pdfium.dll"),
            Path.Combine(currentDir, "runtimes", "win-x64", "native", "pdfium.dll")
        };

        _logger.Debug("Поиск pdfium.dll:");
        foreach (string location in locationsToCheck)
        {
            bool exists = File.Exists(location);
            _logger.Debug($"  {location}: {(exists ? "НАЙДЕН" : "не найден")}");
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
        try
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
        catch (Exception ex)
        {
            _logger.Error($"Ошибка рендеринга страницы {pageIndex + 1}: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Освобождает ресурсы, используемые рендерером.
    /// </summary>
    public void Dispose()
    {
        if (!_disposed)
        {
            try
            {
                _pdfDocument?.Dispose();
                _pdfDocument = null;
            }
            catch (Exception ex)
            {
                _logger.Debug($"Ошибка при освобождении PdfDocument: {ex.Message}");
            }

            _disposed = true;

            // Принудительная сборка мусора для освобождения неуправляемых ресурсов Pdfium
            GC.Collect();
            GC.WaitForPendingFinalizers();

            _logger.Debug("Ресурсы рендерера PDF освобождены");
        }
    }
}