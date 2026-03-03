using System;
using System.Xml.Serialization;

namespace DocToScan.Configuration;

/// <summary>
/// Представляет конфигурацию приложения.
/// </summary>
[XmlRoot("Configuration")]
public class Config
{
    /// <summary>
    /// Получает или задает настройки яркости.
    /// </summary>
    [XmlElement("Brightness")]
    public BrightnessSettings Brightness { get; set; } = new();

    /// <summary>
    /// Получает или задает настройки поворота.
    /// </summary>
    [XmlElement("Rotation")]
    public RotationSettings Rotation { get; set; } = new();

    /// <summary>
    /// Получает или задает настройки градаций серого.
    /// </summary>
    [XmlElement("Grayscale")]
    public GrayscaleSettings Grayscale { get; set; } = new();

    /// <summary>
    /// Получает или задает настройки качества изображений.
    /// </summary>
    [XmlElement("ImageQuality")]
    public ImageQualitySettings ImageQuality { get; set; } = new();

    /// <summary>
    /// Настройки размера страницы
    /// </summary>
    [XmlElement("PageSize")]
    public PageSizeSettings PageSize { get; set; } = new();

    /// <summary>
    /// Настройки яркости изображения.
    /// </summary>
    public class BrightnessSettings
    {
        /// <summary>
        /// Получает или задает значение, указывающее, включено ли изменение яркости.
        /// </summary>
        public bool Enable { get; set; } = true;

        /// <summary>
        /// Получает или задает уровень яркости в диапазоне от -255 до 255.
        /// </summary>
        public int Level { get; set; } = 15;
    }

    /// <summary>
    /// Настройки поворота изображения.
    /// </summary>
    public class RotationSettings
    {
        /// <summary>
        /// Получает или задает значение, указывающее, включен ли поворот.
        /// </summary>
        public bool Enable { get; set; } = true;

        /// <summary>
        /// Получает или задает минимальный угол поворота в градусах.
        /// Отрицательные значения означают поворот влево.
        /// </summary>
        public float MinAngle { get; set; } = -3;

        /// <summary>
        /// Получает или задает максимальный угол поворота в градусах.
        /// Положительные значения означают поворот вправо.
        /// </summary>
        public float MaxAngle { get; set; } = 3;
    }

    /// <summary>
    /// Настройки преобразования в градации серого.
    /// </summary>
    public class GrayscaleSettings
    {
        /// <summary>
        /// Получает или задает значение, указывающее, включено ли преобразование в градации серого.
        /// </summary>
        public bool Enable { get; set; } = true;
    }

    /// <summary>
    /// Настройки качества изображений.
    /// </summary>
    public class ImageQualitySettings
    {
        /// <summary>
        /// Получает или задает разрешение в DPI (точек на дюйм).
        /// </summary>
        public int Dpi { get; set; } = 150;

        /// <summary>
        /// Получает или задает качество JPEG сжатия от 1 (низкое) до 100 (максимальное).
        /// </summary>
        public int JpegCompression { get; set; } = 85;
    }

    /// <summary>
    /// Настройки размера страницы.
    /// </summary>
    public class PageSizeSettings
    {
        /// <summary>
        /// Включить принудительное изменение размера страницы
        /// </summary>
        public bool Enable { get; set; } = true;

        /// <summary>
        /// Ширина страницы в миллиметрах (по умолчанию A4 = 210 мм)
        /// </summary>
        public float WidthMm { get; set; } = 210f;

        /// <summary>
        /// Высота страницы в миллиметрах (по умолчанию A4 = 297 мм)
        /// </summary>
        public float HeightMm { get; set; } = 297f;

        /// <summary>
        /// Целевой формат страницы (A4, A3, Letter, Legal)
        /// </summary>
        [XmlIgnore]
        public PageFormat Format
        {
            get => GetFormatFromSize();
            set => SetSizeFromFormat(value);
        }

        private PageFormat GetFormatFromSize()
        {
            if (Math.Abs(WidthMm - 210) < 1 && Math.Abs(HeightMm - 297) < 1) return PageFormat.A4;
            if (Math.Abs(WidthMm - 297) < 1 && Math.Abs(HeightMm - 420) < 1) return PageFormat.A3;
            if (Math.Abs(WidthMm - 216) < 1 && Math.Abs(HeightMm - 279) < 1) return PageFormat.Letter;
            if (Math.Abs(WidthMm - 216) < 1 && Math.Abs(HeightMm - 356) < 1) return PageFormat.Legal;
            return PageFormat.Custom;
        }

        private void SetSizeFromFormat(PageFormat format)
        {
            switch (format)
            {
                case PageFormat.A4:
                    WidthMm = 210;
                    HeightMm = 297;
                    break;
                case PageFormat.A3:
                    WidthMm = 297;
                    HeightMm = 420;
                    break;
                case PageFormat.Letter:
                    WidthMm = 216;
                    HeightMm = 279;
                    break;
                case PageFormat.Legal:
                    WidthMm = 216;
                    HeightMm = 356;
                    break;
                    // Для Custom оставляем текущие значения
            }
        }
    }

    /// <summary>
    /// Поддерживаемые форматы страниц
    /// </summary>
    public enum PageFormat
    {
        Custom,
        A4,
        A3,
        Letter,
        Legal
    }
}