using SkiaSharp;

namespace DocToScan.ImageProcessing.Effects;

/// <summary>
/// Определяет интерфейс для эффектов обработки изображений.
/// </summary>
public interface IEffect
{
    /// <summary>
    /// Получает название эффекта.
    /// </summary>
    string Name { get; }

    /// <summary>
    /// Применяет эффект к поверхности изображения.
    /// </summary>
    /// <param name="surface">Поверхность для применения эффекта.</param>
    void Apply(SKSurface surface);
}