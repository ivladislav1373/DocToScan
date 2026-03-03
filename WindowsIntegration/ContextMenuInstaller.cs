using System;
using System.IO;
using System.Runtime.InteropServices;
using DocToScan.Logging;
using Microsoft.Win32;

namespace DocToScan.WindowsIntegration;

/// <summary>
/// Управляет интеграцией программы в контекстное меню Проводника Windows.
/// </summary>
public class ContextMenuInstaller
{
    private readonly ILogger _logger;
    private const string MenuCommand = "DocToScan";
    private const string MenuText = "Создать скан-копию";

    /// <summary>
    /// Инициализирует новый экземпляр класса ContextMenuInstaller.
    /// </summary>
    /// <param name="logger">Логгер для записи сообщений.</param>
    public ContextMenuInstaller(ILogger logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <summary>
    /// Устанавливает программу в контекстное меню для PDF и DOCX файлов.
    /// </summary>
    public void Install()
    {
        try
        {
            string exePath = GetExecutablePath();

            _logger.Info("Установка интеграции в контекстное меню...");
            _logger.Debug($"Путь к программе: {exePath}");

            InstallForExtension(".pdf", exePath);
            InstallForExtension(".docx", exePath);
            AddToSendTo(exePath);
            NotifyExplorer();

            _logger.Success("✓ Интеграция успешно установлена!");
            _logger.Info("Теперь вы можете нажать правой кнопкой мыши на PDF или DOCX файл");
            _logger.Info("и выбрать 'Создать скан-копию'");
        }
        catch (UnauthorizedAccessException)
        {
            _logger.Error("✗ Ошибка: Недостаточно прав для изменения реестра.");
            _logger.Info("Запустите программу от имени администратора.");
        }
        catch (Exception ex)
        {
            _logger.Error($"✗ Ошибка установки: {ex.Message}");
            _logger.Debug(ex.StackTrace);
        }
    }

    /// <summary>
    /// Удаляет программу из контекстного меню.
    /// </summary>
    public void Uninstall()
    {
        try
        {
            _logger.Info("Удаление интеграции из контекстного меню...");

            UninstallForExtension(".pdf");
            UninstallForExtension(".docx");
            RemoveFromSendTo();
            NotifyExplorer();

            _logger.Success("✓ Интеграция успешно удалена!");
        }
        catch (UnauthorizedAccessException)
        {
            _logger.Error("✗ Ошибка: Недостаточно прав для изменения реестра.");
            _logger.Info("Запустите программу от имени администратора.");
        }
        catch (Exception ex)
        {
            _logger.Error($"✗ Ошибка удаления: {ex.Message}");
            _logger.Debug(ex.StackTrace);
        }
    }

    /// <summary>
    /// Получает путь к исполняемому файлу программы.
    /// </summary>
    private string GetExecutablePath()
    {
        string exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;

        // Если программа запущена как "dotnet", ищем настоящий exe
        if (exePath.EndsWith("dotnet.exe", StringComparison.OrdinalIgnoreCase))
        {
            exePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DocToScan.exe");
        }

        return exePath;
    }

    /// <summary>
    /// Устанавливает интеграцию для конкретного расширения.
    /// </summary>
    private void InstallForExtension(string extension, string exePath)
    {
        string command = $"\"{exePath}\" \"%1\"";

        using (RegistryKey key = Registry.ClassesRoot.CreateSubKey($"{extension}\\shell\\{MenuCommand}"))
        {
            key.SetValue("", MenuText);
            key.SetValue("Icon", $"\"{exePath}\",0");
        }

        using (RegistryKey key = Registry.ClassesRoot.CreateSubKey($"{extension}\\shell\\{MenuCommand}\\command"))
        {
            key.SetValue("", command);
        }

        _logger.Debug($"Установлено для {extension}");
    }

    /// <summary>
    /// Удаляет интеграцию для конкретного расширения.
    /// </summary>
    private void UninstallForExtension(string extension)
    {
        try
        {
            Registry.ClassesRoot.DeleteSubKeyTree($"{extension}\\shell\\{MenuCommand}", false);
            _logger.Debug($"Удалено для {extension}");
        }
        catch (ArgumentException)
        {
            // Ключ не существует - игнорируем
        }
    }

    /// <summary>
    /// Добавляет программу в меню "Отправить".
    /// </summary>
    private void AddToSendTo(string exePath)
    {
        string sendToPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.SendTo),
            "DocToScan - Создать скан-копию.lnk");

        CreateShortcut(exePath, sendToPath);
        _logger.Debug("Добавлено в меню 'Отправить'");
    }

    /// <summary>
    /// Удаляет программу из меню "Отправить".
    /// </summary>
    private void RemoveFromSendTo()
    {
        string sendToPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.SendTo),
            "DocToScan - Создать скан-копию.lnk");

        if (File.Exists(sendToPath))
        {
            File.Delete(sendToPath);
            _logger.Debug("Удалено из меню 'Отправить'");
        }
    }

    /// <summary>
    /// Создает ярлык для программы.
    /// </summary>
    private void CreateShortcut(string targetPath, string shortcutPath)
    {
        Type t = Type.GetTypeFromCLSID(new Guid("72C24DD5-D70A-438B-8A42-98424B88AFB8")); // Windows Script Host Shell Object
        dynamic shell = Activator.CreateInstance(t);
        try
        {
            var shortcut = shell.CreateShortcut(shortcutPath);
            try
            {
                shortcut.TargetPath = targetPath;
                shortcut.WorkingDirectory = Path.GetDirectoryName(targetPath);
                shortcut.Description = "Создать скан-копию документа";
                shortcut.Save();
            }
            finally
            {
                Marshal.ReleaseComObject(shortcut);
            }
        }
        finally
        {
            Marshal.ReleaseComObject(shell);
        }
    }

    /// <summary>
    /// Уведомляет Проводник Windows об изменениях в реестре.
    /// </summary>
    private void NotifyExplorer()
    {
        try
        {
            // Отправляем сообщение об обновлении всем окнам
            const int HWND_BROADCAST = 0xffff;
            const int WM_SETTINGCHANGE = 0x001a;

            NativeMethods.SendMessageTimeout(
                new IntPtr(HWND_BROADCAST),
                WM_SETTINGCHANGE,
                IntPtr.Zero,
                IntPtr.Zero,
                0,
                1000,
                out _);

            _logger.Debug("Проводник уведомлен об изменениях");
        }
        catch (Exception ex)
        {
            _logger.Debug($"Не удалось уведомить Проводник: {ex.Message}");
        }
    }

    /// <summary>
    /// Внутренний класс для P/Invoke вызовов.
    /// </summary>
    private static class NativeMethods
    {
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessageTimeout(
            IntPtr hWnd,
            int Msg,
            IntPtr wParam,
            IntPtr lParam,
            uint fuFlags,
            uint uTimeout,
            out IntPtr lpdwResult);
    }
}