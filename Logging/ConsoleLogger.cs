using System;
using System.Diagnostics;

namespace DocToScan.Logging;

/// <summary>
/// Реализация логгера, выводящего сообщения в консоль с цветовым форматированием.
/// </summary>
public class ConsoleLogger : ILogger
{
    private readonly bool _isDebugMode;

    /// <summary>
    /// Инициализирует новый экземпляр класса ConsoleLogger.
    /// </summary>
    public ConsoleLogger()
    {
#if DEBUG
        _isDebugMode = true;
#endif
    }

    /// <summary>
    /// Записывает информационное сообщение белым цветом.
    /// </summary>
    public void Info(string message)
    {
        Console.WriteLine(message);
    }

    /// <summary>
    /// Записывает сообщение об ошибке красным цветом.
    /// </summary>
    public void Error(string message)
    {
        var previousColor = Console.ForegroundColor;
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine(message);
        Console.ForegroundColor = previousColor;
    }

    /// <summary>
    /// Записывает предупреждение желтым цветом.
    /// </summary>
    public void Warning(string message)
    {
        var previousColor = Console.ForegroundColor;
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine(message);
        Console.ForegroundColor = previousColor;
    }

    /// <summary>
    /// Записывает сообщение об успехе зеленым цветом.
    /// </summary>
    public void Success(string message)
    {
        var previousColor = Console.ForegroundColor;
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine(message);
        Console.ForegroundColor = previousColor;
    }

    /// <summary>
    /// Записывает отладочное сообщение серым цветом только в Debug режиме.
    /// </summary>
    public void Debug(string message)
    {
        if (!_isDebugMode) return;

        var previousColor = Console.ForegroundColor;
        Console.ForegroundColor = ConsoleColor.DarkGray;
        Console.WriteLine($"[DEBUG] {message}");
        Console.ForegroundColor = previousColor;
    }

    /// <summary>
    /// Записывает разделитель из 55 символов.
    /// </summary>
    public void Separator()
    {
        Console.WriteLine(new string('─', 55));
    }
}