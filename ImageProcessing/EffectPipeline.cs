using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using DocToScan.Configuration;
using DocToScan.ImageProcessing.Effects;
using DocToScan.Logging;
using SkiaSharp;

namespace DocToScan.ImageProcessing;

/// <summary>
/// Представляет конвейер обработки изображений с последовательным применением эффектов
/// </summary>
public class EffectPipeline
{
    private readonly List<IEffect> _effects = new();
    private readonly ILogger _logger;
    private readonly Random _random = new();
    private readonly Config _config;
    private float _appliedRotationAngle = 0;

    /// <summary>
    /// Инициализирует новый экземпляр класса EffectPipeline на основе конфигурации
    /// </summary>
    public EffectPipeline(Config config, ILogger logger)
    {
        _config = config ?? throw new ArgumentNullException(nameof(config));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        InitializeEffects();
    }

    /// <summary>
    /// Инициализирует эффекты согласно конфигурации
    /// </summary>
    private void InitializeEffects()
    {
        if (_config.Rotation.Enable)
        {
            _appliedRotationAngle = (float)(_random.NextDouble() *
                (_config.Rotation.MaxAngle - _config.Rotation.MinAngle) +
                _config.Rotation.MinAngle);

            _effects.Add(new RotationEffect(_appliedRotationAngle));
            _logger.Debug($"Добавлен эффект поворота: {_appliedRotationAngle:F2}°");
        }

        if (_config.Grayscale.Enable)
        {
            _effects.Add(new GrayscaleEffect());
            _logger.Debug("Добавлен эффект градаций серого");
        }

        if (_config.Brightness.Enable && _config.Brightness.Level != 0)
        {
            _effects.Add(new BrightnessEffect(_config.Brightness.Level));
            _logger.Debug($"Добавлен эффект яркости: {_config.Brightness.Level}");
        }

        if (_effects.Count == 0)
        {
            _logger.Debug("Эффекты не применяются");
        }
    }

    /// <summary>
    /// Обрабатывает изображение, применяя все эффекты и обрезку до нужного размера
    /// </summary>
    public Image Process(Image inputImage)
    {
        if (inputImage == null)
        {
            throw new ArgumentNullException(nameof(inputImage));
        }

        // Шаг 1: Применяем эффекты (поворот, яркость, градации серого)
        Bitmap processedImage = ApplyEffects(inputImage);

        // Шаг 2: Обрезаем до нужного размера, если включено
        if (_config.PageSize.Enable)
        {
            _logger.Debug($"Обрезка до размера: {_config.PageSize.WidthMm} x {_config.PageSize.HeightMm} мм");

            processedImage = ImageCropper.CropToSize(
                processedImage,
                _config.PageSize.WidthMm,
                _config.PageSize.HeightMm,
                _config.ImageQuality.Dpi
            );

            _logger.Debug($"Обрезка завершена, новый размер: {processedImage.Width} x {processedImage.Height} пикселей");
        }

        return processedImage;
    }

    /// <summary>
    /// Применяет эффекты к изображению
    /// </summary>
    private Bitmap ApplyEffects(Image inputImage)
    {
        if (_effects.Count == 0)
        {
            return new Bitmap(inputImage);
        }

        // Конвертируем в SKBitmap для обработки
        using var skBitmap = ConvertToSkBitmap(inputImage);
        using var surface = SKSurface.Create(new SKImageInfo(skBitmap.Width, skBitmap.Height));
        var canvas = surface.Canvas;

        canvas.DrawBitmap(skBitmap, 0, 0);

        foreach (var effect in _effects)
        {
            effect.Apply(surface);
            _logger.Debug($"  Применен эффект: {effect.Name}");
        }

        return ConvertToBitmap(surface);
    }

    private SKBitmap ConvertToSkBitmap(Image image)
    {
        using var ms = new MemoryStream();
        image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
        ms.Position = 0;
        return SKBitmap.Decode(ms);
    }

    private Bitmap ConvertToBitmap(SKSurface surface)
    {
        using var skImage = surface.Snapshot();
        using var data = skImage.Encode(SKEncodedImageFormat.Png, 100);
        using var ms = new MemoryStream(data.ToArray());
        return new Bitmap(ms);
    }
}