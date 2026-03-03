using System;
using SkiaSharp;

namespace DocToScan.ImageProcessing.Effects;

/// <summary>
/// Эффект поворота изображения на случайный угол.
/// </summary>
public class RotationEffect : IEffect
{
    private readonly float _angle;

    /// <summary>
    /// Получает название эффекта.
    /// </summary>
    public string Name => "Rotation";

    /// <summary>
    /// Инициализирует новый экземпляр класса RotationEffect.
    /// </summary>
    /// <param name="minAngle">Минимальный угол поворота.</param>
    /// <param name="maxAngle">Максимальный угол поворота.</param>
    public RotationEffect(float minAngle, float maxAngle)
    {
        // Генерируем угол в заданном диапазоне
        _angle = (float)(new Random().NextDouble() * (maxAngle - minAngle) + minAngle);
    }

    /// <summary>
    /// Инициализирует новый экземпляр класса RotationEffect.
    /// </summary>
    /// <param name="angle">Угол поворота.</param>
    public RotationEffect(float angle)
    {
        // Генерируем угол в заданном диапазоне
        _angle = angle;
    }

    /// <summary>
    /// Применяет эффект поворота к поверхности.
    /// </summary>
    public void Apply(SKSurface surface)
    {
        if (Math.Abs(_angle) < 0.01f) return;

        var canvas = surface.Canvas;
        var width = surface.Canvas.LocalClipBounds.Width;
        var height = surface.Canvas.LocalClipBounds.Height;

        // Поворачиваем относительно центра изображения
        canvas.Translate(width / 2f, height / 2f);
        canvas.RotateDegrees(_angle);
        canvas.Translate(-width / 2f, -height / 2f);
    }
}