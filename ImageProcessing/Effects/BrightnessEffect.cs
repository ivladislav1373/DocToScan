using SkiaSharp;

namespace DocToScan.ImageProcessing.Effects;

/// <summary>
/// Эффект изменения яркости изображения.
/// </summary>
public class BrightnessEffect : IEffect
{
    private readonly int _level;

    /// <summary>
    /// Получает название эффекта.
    /// </summary>
    public string Name => "Brightness";

    /// <summary>
    /// Инициализирует новый экземпляр класса BrightnessEffect.
    /// </summary>
    /// <param name="level">Уровень яркости от -255 до 255.</param>
    public BrightnessEffect(int level)
    {
        _level = level;
    }

    /// <summary>
    /// Применяет эффект яркости к поверхности.
    /// </summary>
    /// <param name="surface">Поверхность для применения эффекта.</param>
    public void Apply(SKSurface surface)
    {
        if (_level == 0) return;

        float brightness = _level / 255f;

        using var paint = new SKPaint
        {
            ColorFilter = SKColorFilter.CreateColorMatrix(new float[]
            {
                1, 0, 0, 0, brightness,
                0, 1, 0, 0, brightness,
                0, 0, 1, 0, brightness,
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