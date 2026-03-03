namespace DocToScan.Logging;

/// <summary>
/// Определяет интерфейс для логирования сообщений.
/// </summary>
public interface ILogger
{
    /// <summary>
    /// Записывает информационное сообщение.
    /// </summary>
    /// <param name="message">Текст сообщения.</param>
    void Info(string message);

    /// <summary>
    /// Записывает сообщение об ошибке.
    /// </summary>
    /// <param name="message">Текст сообщения.</param>
    void Error(string message);

    /// <summary>
    /// Записывает предупреждение.
    /// </summary>
    /// <param name="message">Текст сообщения.</param>
    void Warning(string message);

    /// <summary>
    /// Записывает сообщение об успешной операции.
    /// </summary>
    /// <param name="message">Текст сообщения.</param>
    void Success(string message);

    /// <summary>
    /// Записывает отладочное сообщение (только в Debug режиме).
    /// </summary>
    /// <param name="message">Текст сообщения.</param>
    void Debug(string message);

    /// <summary>
    /// Записывает разделитель для улучшения читаемости лога.
    /// </summary>
    void Separator();
}