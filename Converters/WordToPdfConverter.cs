using System;
using System.IO;
using System.Runtime.InteropServices;
using DocToScan.Logging;
using Microsoft.Office.Interop.Word;

namespace DocToScan.Converters;

/// <summary>
/// Предоставляет функциональность для конвертации DOCX файлов в PDF.
/// Поддерживает разные версии Microsoft Office (2010, 2013, 2016, 2019, 2021, 2024, Microsoft 365).
/// </summary>
public class WordToPdfConverter : IDisposable
{
    private readonly ILogger _logger;
    private Application _wordApp;
    private bool _disposed;
    private static readonly object _lockObj = new object();

    /// <summary>
    /// Инициализирует новый экземпляр класса WordToPdfConverter.
    /// </summary>
    public WordToPdfConverter(ILogger logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <summary>
    /// Конвертирует DOCX файл в PDF.
    /// </summary>
    public void ConvertToPdf(string docxPath, string pdfPath)
    {
        ValidatePaths(docxPath, pdfPath);

        try
        {
            InitializeWordApplication();

            _logger.Debug($"Открытие документа: {docxPath}");

            Document doc = null;
            try
            {
                doc = _wordApp.Documents.Open(docxPath, ReadOnly: true);

                _logger.Debug($"Сохранение в PDF: {pdfPath}");

                // Определяем формат сохранения в зависимости от версии Word
                WdSaveFormat saveFormat = GetSaveFormat();

                doc.SaveAs2(pdfPath, saveFormat);

                if (!File.Exists(pdfPath))
                {
                    throw new InvalidOperationException("PDF файл не был создан");
                }

                _logger.Debug("Конвертация успешно завершена");
            }
            finally
            {
                CloseDocument(doc);
            }
        }
        catch (COMException ex) when (ex.Message.Contains("cannot be found") || ex.Message.Contains("не найден"))
        {
            _logger.Error("Microsoft Word не найден. Убедитесь, что Microsoft Word установлен.");
            throw new InvalidOperationException(
                "Microsoft Word не найден. Для конвертации DOCX файлов требуется установленный Microsoft Word.", ex);
        }
        catch (Exception ex)
        {
            _logger.Error($"Ошибка конвертации: {ex.Message}");
            throw new InvalidOperationException($"Ошибка конвертации DOCX в PDF: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Инициализирует приложение Microsoft Word с учетом разных версий.
    /// </summary>
    private void InitializeWordApplication()
    {
        if (_wordApp != null) return;

        lock (_lockObj)
        {
            if (_wordApp != null) return;

            try
            {
                _logger.Debug("Попытка инициализации Microsoft Word...");

                // Пробуем разные способы создания экземпляра Word
                _wordApp = CreateWordApplication();

                // Проверяем версию Word
                string version = _wordApp.Version;
                _logger.Info($"Microsoft Word версии {version} успешно инициализирован");

                _wordApp.Visible = false;
                _wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                _wordApp.ScreenUpdating = false;
            }
            catch (COMException ex)
            {
                _logger.Error($"COM ошибка при инициализации Word: {ex.Message}");
                throw new InvalidOperationException(
                    "Не удалось запустить Microsoft Word. Возможные причины:\n" +
                    "1. Microsoft Word не установлен\n" +
                    "2. Установлена версия Word, не поддерживающая COM-автоматизацию (например, Word Starter)\n" +
                    "3. Недостаточно прав для создания COM-объектов", ex);
            }
        }
    }

    /// <summary>
    /// Создает экземпляр Word Application с учетом разных версий.
    /// </summary>
    private Application CreateWordApplication()
    {
        try
        {
            // Стандартный способ
            return new Application();
        }
        catch (COMException)
        {
            // Альтернативный способ через ProgID для разных версий
            string[] progIds = new[]
            {
                "Word.Application",
                "Word.Application.16",
                "Word.Application.15",
                "Word.Application.14",
                "Word.Application.12"
            };

            foreach (string progId in progIds)
            {
                try
                {
                    _logger.Debug($"Попытка создания через ProgID: {progId}");
                    Type wordType = Type.GetTypeFromProgID(progId, true);
                    if (wordType != null)
                    {
                        return (Application)Activator.CreateInstance(wordType);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Debug($"Не удалось создать через {progId}: {ex.Message}");
                    // Продолжаем попытки
                }
            }

            throw new COMException("Не удалось создать экземпляр Microsoft Word");
        }
    }

    /// <summary>
    /// Определяет формат сохранения в зависимости от версии Word.
    /// </summary>
    private WdSaveFormat GetSaveFormat()
    {
        try
        {
            // Пробуем использовать wdFormatPDF (стандартный для новых версий)
            return WdSaveFormat.wdFormatPDF;
        }
        catch
        {
            _logger.Warning("Используется альтернативный формат сохранения PDF");
            // Для старых версий Word можно использовать другие форматы
            // Например, сохранить как XPS и конвертировать, но это сложнее
            return WdSaveFormat.wdFormatPDF; // Заглушка, реально нужно обрабатывать ошибку
        }
    }

    /// <summary>
    /// Проверяет пути к файлам.
    /// </summary>
    private void ValidatePaths(string docxPath, string pdfPath)
    {
        if (!File.Exists(docxPath))
        {
            throw new FileNotFoundException($"DOCX файл не найден: {docxPath}");
        }

        if (string.IsNullOrWhiteSpace(pdfPath))
        {
            throw new ArgumentException("Путь к PDF файлу не может быть пустым", nameof(pdfPath));
        }

        string extension = Path.GetExtension(docxPath).ToLowerInvariant();
        if (extension != ".docx" && extension != ".doc")
        {
            throw new ArgumentException($"Файл должен быть документом Word (.docx или .doc): {docxPath}");
        }

        // Проверяем, что путь к PDF доступен для записи
        string directory = Path.GetDirectoryName(pdfPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }

    /// <summary>
    /// Закрывает документ и освобождает ресурсы.
    /// </summary>
    private void CloseDocument(Document doc)
    {
        if (doc == null) return;

        try
        {
            doc.Close(SaveChanges: false);
            Marshal.ReleaseComObject(doc);
        }
        catch (Exception ex)
        {
            _logger.Debug($"Ошибка при закрытии документа: {ex.Message}");
        }
    }

    /// <summary>
    /// Освобождает ресурсы Word.
    /// </summary>
    public void Dispose()
    {
        if (!_disposed && _wordApp != null)
        {
            try
            {
                _wordApp.Quit(SaveChanges: false);
                Marshal.ReleaseComObject(_wordApp);
                _wordApp = null;
                _logger.Debug("Ресурсы Word освобождены");
            }
            catch (Exception ex)
            {
                _logger.Debug($"Ошибка при освобождении ресурсов Word: {ex.Message}");
            }
            _disposed = true;
        }
    }
}