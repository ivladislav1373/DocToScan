using System;
using System.Collections.Generic;
using System.IO;
using DocToScan.Logging;

namespace DocToScan.FileSystem;

/// <summary>
/// Управляет временными файлами, обеспечивая их удаление после использования.
/// </summary>
public class TempFileManager : IDisposable
{
    private readonly ILogger _logger;
    private readonly List<string> _tempFiles = new();
    private readonly string _tempDirectory;
    private bool _disposed;

    /// <summary>
    /// Инициализирует новый экземпляр класса TempFileManager.
    /// </summary>
    /// <param name="logger">Логгер для записи сообщений.</param>
    public TempFileManager(ILogger logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _tempDirectory = Path.Combine(Path.GetTempPath(), "DocToScan");

        EnsureTempDirectoryExists();
    }

    /// <summary>
    /// Создает временный файл с указанным расширением.
    /// </summary>
    /// <param name="extension">Расширение файла (без точки).</param>
    /// <returns>Путь к созданному временному файлу.</returns>
    public string CreateTempFile(string extension)
    {
        string fileName = $"DocToScan_{Guid.NewGuid():N}.{extension}";
        string filePath = Path.Combine(_tempDirectory, fileName);

        _tempFiles.Add(filePath);
        _logger.Debug($"Создан временный файл: {fileName}");

        return filePath;
    }

    /// <summary>
    /// Регистрирует существующий файл для последующего удаления.
    /// </summary>
    /// <param name="filePath">Путь к файлу.</param>
    public void RegisterFile(string filePath)
    {
        if (!string.IsNullOrEmpty(filePath) && !_tempFiles.Contains(filePath))
        {
            _tempFiles.Add(filePath);
            _logger.Debug($"Зарегистрирован файл для удаления: {Path.GetFileName(filePath)}");
        }
    }

    /// <summary>
    /// Немедленно удаляет все зарегистрированные временные файлы.
    /// </summary>
    public void Cleanup()
    {
        foreach (string filePath in _tempFiles)
        {
            DeleteFile(filePath);
        }

        _tempFiles.Clear();
        DeleteTempDirectoryIfEmpty();
    }

    /// <summary>
    /// Удаляет конкретный файл.
    /// </summary>
    private void DeleteFile(string filePath)
    {
        try
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
                _logger.Debug($"Удален временный файл: {Path.GetFileName(filePath)}");
            }
        }
        catch (Exception ex)
        {
            _logger.Debug($"Не удалось удалить временный файл {filePath}: {ex.Message}");
        }
    }

    /// <summary>
    /// Удаляет временную директорию, если она пуста.
    /// </summary>
    private void DeleteTempDirectoryIfEmpty()
    {
        try
        {
            if (Directory.Exists(_tempDirectory) &&
                Directory.GetFiles(_tempDirectory).Length == 0)
            {
                Directory.Delete(_tempDirectory);
                _logger.Debug("Временная директория удалена");
            }
        }
        catch (Exception ex)
        {
            _logger.Debug($"Не удалось удалить временную директорию: {ex.Message}");
        }
    }

    /// <summary>
    /// Создает временную директорию, если она не существует.
    /// </summary>
    private void EnsureTempDirectoryExists()
    {
        if (!Directory.Exists(_tempDirectory))
        {
            Directory.CreateDirectory(_tempDirectory);
            _logger.Debug($"Создана временная директория: {_tempDirectory}");
        }
    }

    /// <summary>
    /// Освобождает ресурсы и удаляет временные файлы.
    /// </summary>
    public void Dispose()
    {
        if (!_disposed)
        {
            Cleanup();
            _disposed = true;
        }
    }
}