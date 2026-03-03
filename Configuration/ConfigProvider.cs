using System;
using System.IO;
using System.Xml.Serialization;
using DocToScan.Logging;

namespace DocToScan.Configuration;

/// <summary>
/// Предоставляет методы для загрузки и сохранения конфигурации приложения.
/// </summary>
public static class ConfigProvider
{
    private static readonly ILogger _logger = new ConsoleLogger();

    /// <summary>
    /// Загружает конфигурацию из XML файла.
    /// </summary>
    /// <param name="configPath">Путь к файлу конфигурации.</param>
    /// <returns>Объект конфигурации или null, если файл не найден.</returns>
    public static Config Load(string configPath)
    {
        try
        {
            if (!File.Exists(configPath))
            {
                _logger.Debug($"Файл конфигурации не найден: {configPath}");
                return null;
            }

            var serializer = new XmlSerializer(typeof(Config));
            using var reader = new StreamReader(configPath);
            var config = (Config)serializer.Deserialize(reader);

            ValidateConfig(config);
            _logger.Debug($"Конфигурация успешно загружена из {configPath}");

            return config;
        }
        catch (Exception ex)
        {
            _logger.Warning($"Ошибка загрузки конфигурации: {ex.Message}. Используются значения по умолчанию.");
            return null;
        }
    }

    /// <summary>
    /// Сохраняет конфигурацию в XML файл.
    /// </summary>
    /// <param name="config">Объект конфигурации для сохранения.</param>
    /// <param name="configPath">Путь к файлу конфигурации.</param>
    public static void Save(Config config, string configPath)
    {
        try
        {
            var serializer = new XmlSerializer(typeof(Config));
            using var writer = new StreamWriter(configPath);
            serializer.Serialize(writer, config);

            _logger.Debug($"Конфигурация сохранена в {configPath}");
        }
        catch (Exception ex)
        {
            _logger.Error($"Ошибка сохранения конфигурации: {ex.Message}");
        }
    }

    /// <summary>
    /// Проверяет корректность значений конфигурации и корректирует их при необходимости.
    /// </summary>
    /// <param name="config">Объект конфигурации для проверки.</param>
    private static void ValidateConfig(Config config)
    {
        // Проверка яркости
        if (config.Brightness.Level < -255 || config.Brightness.Level > 255)
        {
            _logger.Warning($"Уровень яркости {config.Brightness.Level} вне допустимого диапазона (-255..255). Используется 0.");
            config.Brightness.Level = 0;
        }

        // Проверка углов поворота
        if (config.Rotation.MinAngle > config.Rotation.MaxAngle)
        {
            _logger.Warning("Минимальный угол поворота больше максимального. Значения поменяны местами.");
            (config.Rotation.MinAngle, config.Rotation.MaxAngle) = (config.Rotation.MaxAngle, config.Rotation.MinAngle);
        }

        // Проверка качества JPEG
        if (config.ImageQuality.JpegCompression < 1 || config.ImageQuality.JpegCompression > 100)
        {
            _logger.Warning($"Качество JPEG {config.ImageQuality.JpegCompression} вне допустимого диапазона (1..100). Используется 85.");
            config.ImageQuality.JpegCompression = 85;
        }

        // Проверка DPI
        if (config.ImageQuality.Dpi < 72 || config.ImageQuality.Dpi > 600)
        {
            _logger.Warning($"DPI {config.ImageQuality.Dpi} вне рекомендуемого диапазона (72..600). Используется 150.");
            config.ImageQuality.Dpi = 150;
        }
    }
}