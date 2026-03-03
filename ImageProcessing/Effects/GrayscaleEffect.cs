using SkiaSharp;

namespace DocToScan.ImageProcessing.Effects;

/// <summary>
/// Эффект преобразования изображения в градации серого.
/// </summary>
public class GrayscaleEffect : IEffect
{
    /// <summary>
    /// Получает название эффекта.
    /// </summary>
    public string Name => "Grayscale";

    /// <summary>
    /// Применяет эффект градаций серого к поверхности.
    /// </summary>
    /// <param name="surface">Поверхность для применения эффекта.</param>
    public void Apply(SKSurface surface)
    {
        using var paint = new SKPaint
        {
            ColorFilter = SKColorFilter.CreateColorMatrix(new float[]
            {
                0.3f, 0.59f, 0.11f, 0, 0,
                0.3f, 0.59f, 0.11f, 0, 0,
                0.3f, 0.59f, 0.11f, 0, 0,
                0, 0, 0, 1, 0
            }),
            BlendMode = SKBlendMode.SrcOver
        };

        var canvas = surface.Canvas;
        using var snapshot = surface.Snapshot();
        canvas.Clear();
        canvas.DrawImage(snapshot, 0, 0, paint);
    }
}