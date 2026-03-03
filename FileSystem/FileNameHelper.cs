using System.IO;

namespace DocToScan.FileSystem;

/// <summary>
/// Предоставляет вспомогательные методы для работы с именами файлов.
/// </summary>
public static class FileNameHelper
{
    /// <summary>
    /// Генерирует уникальное имя файла, добавляя суффикс и индекс при необходимости.
    /// </summary>
    /// <param name="originalPath">Путь к исходному файлу.</param>
    /// <param name="suffix">Суффикс для добавления к имени файла.</param>
    /// <returns>Уникальный путь к файлу, который не существует.</returns>
    /// <example>
    /// Если файл "doc.pdf" существует, будет создано "doc (скан).pdf"
    /// Если такой тоже существует, то "doc (скан) [1].pdf" и т.д.
    /// </example>
    public static string GetUniqueOutputPath(string originalPath, string suffix = "скан")
    {
        string directory = Path.GetDirectoryName(originalPath);
        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(originalPath);
        string extension = ".pdf"; // Всегда сохраняем как PDF

        // Базовое имя без индекса
        string baseName = $"{fileNameWithoutExt} ({suffix})";
        string outputPath = Path.Combine(directory, baseName + extension);

        // Если файл не существует, возвращаем его
        if (!File.Exists(outputPath))
            return outputPath;

        // Иначе добавляем индекс
        int counter = 1;
        while (true)
        {
            string indexedName = $"{fileNameWithoutExt} ({suffix}) [{counter}]{extension}";
            outputPath = Path.Combine(directory, indexedName);

            if (!File.Exists(outputPath))
                break;

            counter++;
        }

        return outputPath;
    }

    /// <summary>
    /// Проверяет, является ли файл изображением по расширению.
    /// </summary>
    /// <param name="filePath">Путь к файлу.</param>
    /// <returns>true, если файл имеет расширение изображения.</returns>
    public static bool IsImageFile(string filePath)
    {
        string extension = Path.GetExtension(filePath).ToLowerInvariant();
        return extension == ".jpg" || extension == ".jpeg" ||
               extension == ".png" || extension == ".bmp" ||
               extension == ".tiff";
    }

    /// <summary>
    /// Получает безопасное имя файла, удаляя недопустимые символы.
    /// </summary>
    /// <param name="fileName">Исходное имя файла.</param>
    /// <returns>Имя файла с удаленными недопустимыми символами.</returns>
    public static string GetSafeFileName(string fileName)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            fileName = fileName.Replace(c, '_');
        }
        return fileName;
    }
}