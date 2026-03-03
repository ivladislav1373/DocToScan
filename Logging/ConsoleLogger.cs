using System;

namespace DocToScan.Logging;

/// <summary>
/// Реализация логгера, выводящего сообщения в консоль.
/// </summary>
public class ConsoleLogger : ILogger
{
    private readonly bool _isDebugMode;

    /// <summary>
    /// Инициализирует новый экземпляр класса ConsoleLogger.
    /// </summary>
    /// <param name="debugMode">Включить отладочный режим.</param>
    public ConsoleLogger(bool debugMode = false)
    {
        _isDebugMode = debugMode;

        // Настраиваем кодировку консоли для поддержки UTF-8
        try
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
        }
        catch
        {
            // Игнорируем ошибки кодировки
        }
    }

    /// <inheritdoc />
    public void Info(string message)
    {
        Console.ForegroundColor = ConsoleColor.White;
        Console.WriteLine(message);
        Console.ResetColor();
    }

    /// <inheritdoc />
    public void Success(string message)
    {
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine(message);
        Console.ResetColor();
    }

    /// <inheritdoc />
    public void Warning(string message)
    {
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine(message);
        Console.ResetColor();
    }

    /// <inheritdoc />
    public void Error(string message)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine(message);
        Console.ResetColor();
    }

    /// <inheritdoc />
    public void Debug(string message)
    {
        if (_isDebugMode)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine($"[DEBUG] {message}");
            Console.ResetColor();
        }
    }

    /// <inheritdoc />
    public void Separator()
    {
        Console.ForegroundColor = ConsoleColor.DarkGray;
        Console.WriteLine(new string('─', 55));
        Console.ResetColor();
    }
}