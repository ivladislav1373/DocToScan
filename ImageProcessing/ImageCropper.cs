using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using DocToScan.Configuration;

namespace DocToScan.ImageProcessing;

/// <summary>
/// Выполняет обрезку изображения до заданного размера с учетом поворота
/// </summary>
public static class ImageCropper
{
    /// <summary>
    /// Обрезает изображение до целевого размера с учетом поворота
    /// </summary>
    /// <param name="image">Исходное изображение (уже повернутое)</param>
    /// <param name="targetWidthMm">Целевая ширина в миллиметрах</param>
    /// <param name="targetHeightMm">Целевая высота в миллиметрах</param>
    /// <param name="dpi">Разрешение в DPI</param>
    /// <returns>Обрезанное изображение</returns>
    public static Bitmap CropToSize(Bitmap image, float targetWidthMm, float targetHeightMm, int dpi)
    {
        // Конвертируем миллиметры в пиксели с учетом DPI
        int targetWidthPx = MmToPixels(targetWidthMm, dpi);
        int targetHeightPx = MmToPixels(targetHeightMm, dpi);

        // Создаем новое изображение целевого размера
        var result = new Bitmap(targetWidthPx, targetHeightPx);
        result.SetResolution(dpi, dpi);

        using (var g = Graphics.FromImage(result))
        {
            g.Clear(Color.White);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;

            // Вычисляем масштаб для вписывания изображения
            float scaleX = (float)targetWidthPx / image.Width;
            float scaleY = (float)targetHeightPx / image.Height;

            // Используем Max чтобы изображение покрыло всю страницу (будут обрезаны края)
            // или Min чтобы изображение полностью вписалось (будут поля)
            float scale = Math.Max(scaleX, scaleY);

            // Вычисляем размер после масштабирования
            int scaledWidth = (int)(image.Width * scale);
            int scaledHeight = (int)(image.Height * scale);

            // Вычисляем позицию для центрирования
            int x = (targetWidthPx - scaledWidth) / 2;
            int y = (targetHeightPx - scaledHeight) / 2;

            // ВАЖНО: НЕ применяем поворот здесь, так как изображение уже повернуто!
            // Просто рисуем изображение с масштабированием
            g.DrawImage(image, x, y, scaledWidth, scaledHeight);
        }

        return result;
    }

    /// <summary>
    /// Конвертирует миллиметры в пиксели с учетом DPI
    /// </summary>
    private static int MmToPixels(float mm, int dpi)
    {
        // 1 дюйм = 25.4 мм
        // Количество дюймов = mm / 25.4
        // Количество пикселей = дюймы * dpi
        return (int)((mm / 25.4f) * dpi + 0.5f);
    }

    /// <summary>
    /// Конвертирует пиксели в миллиметры с учетом DPI
    /// </summary>
    public static float PixelsToMm(int pixels, int dpi)
    {
        return (pixels * 25.4f) / dpi;
    }
}