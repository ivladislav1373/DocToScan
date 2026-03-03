using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using DocToScan.Configuration;
using DocToScan.Logging;
using DocToScan.Processing;
using DocToScan.WindowsIntegration;

namespace DocToScan;

/// <summary>
/// Главный класс программы, точка входа в приложение.
/// Обрабатывает аргументы командной строки и управляет жизненным циклом приложения.
/// </summary>
internal class Program
{
    private static ILogger _logger;

    /// <summary>
    /// Точка входа в приложение.
    /// </summary>
    /// <param name="args">Аргументы командной строки.</param>
    static void Main(string[] args)
    {
        // Настройка обработки сборок Office Interop
        SetupAssemblyResolver();

        // Настройка консоли для поддержки UTF-8
        SetupConsoleEncoding();

        // Инициализация логгера
        _logger = new ConsoleLogger();

        _logger.Info("DocToScan v1.0 - Преобразование документов в \"скан-копии\"");
        _logger.Separator();

        try
        {
            if (args.Length == 0)
            {
                ShowUsage();
                WaitForExit();
                return;
            }

            // Обработка специальных команд
            string firstArg = args[0].ToLowerInvariant();
            if (IsInstallCommand(firstArg))
            {
                HandleInstallCommand();
                WaitForExit();
                return;
            }

            if (IsUninstallCommand(firstArg))
            {
                HandleUninstallCommand();
                WaitForExit();
                return;
            }

            // Обработка файлов
            ProcessFiles(args);
        }
        catch (Exception ex)
        {
            _logger.Error($"Критическая ошибка: {ex.Message}");
            _logger.Debug(ex.StackTrace);
        }

        WaitForExit();
    }

    /// <summary>
    /// Настраивает обработчик для загрузки сборок Office Interop.
    /// </summary>
    private static void SetupAssemblyResolver()
    {
        AppDomain.CurrentDomain.AssemblyResolve += ResolveOfficeInterop;
    }

    /// <summary>
    /// Настраивает кодировку консоли для правильного отображения русских символов.
    /// </summary>
    private static void SetupConsoleEncoding()
    {
        try
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.InputEncoding = System.Text.Encoding.UTF8;
        }
        catch
        {
            // Игнорируем ошибки кодировки
        }
    }

    /// <summary>
    /// Обработчик загрузки сборок Office Interop.
    /// </summary>
    private static Assembly ResolveOfficeInterop(object sender, ResolveEventArgs args)
    {
        string assemblyName = new AssemblyName(args.Name).Name;

        // Проверяем, нужно ли обрабатывать эту сборку
        if (!IsOfficeInteropAssembly(assemblyName))
        {
            return null;
        }

        _logger.Info($"Загрузка сборки Office Interop: {assemblyName}");

        // Сначала проверяем GAC (самый надежный путь)
        Assembly gacAssembly = LoadFromGac(assemblyName);
        if (gacAssembly != null)
        {
            _logger.Success($"Сборка загружена из GAC: {assemblyName}");
            return gacAssembly;
        }

        // Если не нашли в GAC, ищем в других местах
        Assembly foundAssembly = FindOfficeAssembly(assemblyName);
        if (foundAssembly != null)
        {
            _logger.Success($"Сборка загружена: {assemblyName}");
            return foundAssembly;
        }

        _logger.Error($"Не удалось найти сборку Office Interop: {assemblyName}");
        _logger.Info("Убедитесь, что установлен Microsoft Office и компоненты Primary Interop Assemblies (PIA)");

        return null;
    }

    /// <summary>
    /// Пытается загрузить сборку из GAC.
    /// </summary>
    private static Assembly LoadFromGac(string assemblyName)
    {
        try
        {
            // Специфичные пути для Microsoft.Office.Interop.Word
            if (assemblyName.Equals("Microsoft.Office.Interop.Word", StringComparison.OrdinalIgnoreCase))
            {
                string gacPath = @"C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll";
                if (File.Exists(gacPath))
                {
                    _logger.Debug($"Найдена сборка в GAC: {gacPath}");
                    return Assembly.LoadFrom(gacPath);
                }
            }

            // Для office.dll
            if (assemblyName.Equals("office", StringComparison.OrdinalIgnoreCase))
            {
                string gacPath = @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll";
                if (File.Exists(gacPath))
                {
                    _logger.Debug($"Найдена сборка в GAC: {gacPath}");
                    return Assembly.LoadFrom(gacPath);
                }
            }

            // Общий поиск в GAC
            string gacBasePath = @"C:\Windows\assembly\GAC_MSIL";
            if (Directory.Exists(gacBasePath))
            {
                foreach (string dir in Directory.GetDirectories(gacBasePath))
                {
                    string dirName = Path.GetFileName(dir);
                    if (dirName.StartsWith(assemblyName, StringComparison.OrdinalIgnoreCase) ||
                        (assemblyName.Equals("office") && dirName.StartsWith("office", StringComparison.OrdinalIgnoreCase)))
                    {
                        foreach (string versionDir in Directory.GetDirectories(dir))
                        {
                            foreach (string dllPath in Directory.GetFiles(versionDir, "*.dll"))
                            {
                                _logger.Debug($"Найдена возможная сборка в GAC: {dllPath}");
                                return Assembly.LoadFrom(dllPath);
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.Debug($"Ошибка при загрузке из GAC: {ex.Message}");
        }

        return null;
    }

    /// <summary>
    /// Проверяет, является ли имя сборки сборкой Office Interop.
    /// </summary>
    private static bool IsOfficeInteropAssembly(string assemblyName)
    {
        return assemblyName.Contains("Office.Interop.Word") ||
               assemblyName.Equals("office", StringComparison.OrdinalIgnoreCase) ||
               assemblyName.Equals("Microsoft.Office.Interop.Word", StringComparison.OrdinalIgnoreCase) ||
               assemblyName.StartsWith("Microsoft.Office.", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Ищет сборку Office по всем возможным путям.
    /// </summary>
    private static Assembly FindOfficeAssembly(string assemblyName)
    {
        string[] searchPaths = new[]
        {
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
            @"C:\Program Files\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\OFFICE16",
            @"C:\Program Files (x86)\Microsoft Office\root\vfs\ProgramFilesCommonX86\Microsoft Shared\OFFICE16",
            @"C:\Program Files\Microsoft Office\Office16",
            @"C:\Program Files (x86)\Microsoft Office\Office16",
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + @"\Microsoft Office\root\vfs\ProgramFilesCommonX64\Microsoft Shared\OFFICE16",
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86) + @"\Microsoft Office\root\vfs\ProgramFilesCommonX86\Microsoft Shared\OFFICE16",
            @"C:\Program Files\Common Files\Microsoft Shared\OFFICE16",
            @"C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16"
        };

        string[] searchPatterns = new[]
        {
            $"{assemblyName}.dll",
            "Microsoft.Office.Interop.Word.dll",
            "office.dll"
        };

        foreach (string basePath in searchPaths)
        {
            if (string.IsNullOrEmpty(basePath) || !Directory.Exists(basePath))
                continue;

            foreach (string pattern in searchPatterns)
            {
                try
                {
                    // Ищем в текущей папке
                    string fullPath = Path.Combine(basePath, pattern);
                    if (File.Exists(fullPath))
                    {
                        _logger.Debug($"Найдена сборка: {fullPath}");
                        return Assembly.LoadFrom(fullPath);
                    }

                    // Ищем в подпапках (только первый уровень)
                    foreach (string subDir in Directory.GetDirectories(basePath))
                    {
                        string subPath = Path.Combine(subDir, pattern);
                        if (File.Exists(subPath))
                        {
                            _logger.Debug($"Найдена сборка: {subPath}");
                            return Assembly.LoadFrom(subPath);
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.Debug($"Ошибка при поиске {pattern} в {basePath}: {ex.Message}");
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Определяет, является ли аргумент командой установки.
    /// </summary>
    private static bool IsInstallCommand(string arg)
    {
        return arg == "/install" || arg == "-install" || arg == "--install";
    }

    /// <summary>
    /// Определяет, является ли аргумент командой удаления.
    /// </summary>
    private static bool IsUninstallCommand(string arg)
    {
        return arg == "/uninstall" || arg == "-uninstall" || arg == "--uninstall";
    }

    /// <summary>
    /// Обрабатывает команду установки в контекстное меню.
    /// </summary>
    private static void HandleInstallCommand()
    {
        try
        {
            _logger.Info("Установка интеграции в контекстное меню...");
            var installer = new ContextMenuInstaller(_logger);
            installer.Install();
            _logger.Success("✓ Интеграция успешно установлена!");
        }
        catch (Exception ex)
        {
            _logger.Error($"✗ Ошибка при установке: {ex.Message}");
        }
    }

    /// <summary>
    /// Обрабатывает команду удаления из контекстного меню.
    /// </summary>
    private static void HandleUninstallCommand()
    {
        try
        {
            _logger.Info("Удаление интеграции из контекстного меню...");
            var installer = new ContextMenuInstaller(_logger);
            installer.Uninstall();
            _logger.Success("✓ Интеграция успешно удалена!");
        }
        catch (Exception ex)
        {
            _logger.Error($"✗ Ошибка при удалении: {ex.Message}");
        }
    }

    /// <summary>
    /// Обрабатывает переданные файлы.
    /// </summary>
    private static void ProcessFiles(string[] args)
    {
        // Загрузка конфигурации
        var config = LoadConfiguration();

        // Получаем список валидных файлов
        var filesToProcess = GetValidFiles(args);

        if (filesToProcess.Count == 0)
        {
            _logger.Warning("Не найдено ни одного существующего файла для обработки.");
            ShowUsage();
            return;
        }

        _logger.Info($"Найдено файлов для обработки: {filesToProcess.Count}");
        _logger.Separator();

        int successCount = 0;
        int totalPages = 0;

        // Обрабатываем каждый файл
        foreach (string filePath in filesToProcess)
        {
            _logger.Info($"Файл: {filePath}");

            try
            {
                var processor = new DocumentProcessor(config, _logger);
                var result = processor.ProcessFile(filePath);

                if (result.Success)
                {
                    successCount++;
                    totalPages += result.PageCount;
                    _logger.Success($"  ✓ Готово: {Path.GetFileName(result.OutputPath)}");

                    // Показываем информацию о промежуточном PDF, если он был создан
                    if (result.HasIntermediatePdf)
                    {
                        _logger.Info($"    └─ Промежуточный PDF: {Path.GetFileName(result.IntermediatePdfPath)}");
                    }
                }
                else
                {
                    _logger.Error($"  ✗ Ошибка: {result.ErrorMessage}");
                }

            }
            catch (Exception ex)
            {
                _logger.Error($"  ✗ Непредвиденная ошибка: {ex.Message}");
                _logger.Debug(ex.StackTrace);
            }

            _logger.Separator();
        }

        _logger.Info($"Обработано: {successCount} файлов, {totalPages} страниц");
    }

    /// <summary>
    /// Загружает конфигурацию приложения.
    /// </summary>
    private static Config LoadConfiguration()
    {
        var config = ConfigProvider.Load("config.xml");
        if (config == null)
        {
            _logger.Warning("Конфигурация не найдена, используются значения по умолчанию");
            config = new Config();

            // Сохраняем конфигурацию по умолчанию для удобства пользователя
            try
            {
                ConfigProvider.Save(config, "config.xml");
                _logger.Info("Создан файл конфигурации по умолчанию: config.xml");
            }
            catch (Exception ex)
            {
                _logger.Debug($"Не удалось сохранить конфигурацию: {ex.Message}");
            }
        }

        _logger.Info($"Конфигурация загружена: DPI={config.ImageQuality.Dpi}, " +
                      $"Поворот={config.Rotation.Enable}, Яркость={config.Brightness.Enable}, Ч/Б={config.Grayscale.Enable}");

        return config;
    }

    /// <summary>
    /// Получает список существующих файлов из аргументов командной строки.
    /// </summary>
    private static List<string> GetValidFiles(string[] args)
    {
        return args
            .Where(arg => !arg.StartsWith("/") && !arg.StartsWith("-"))
            .Where(File.Exists)
            .ToList();
    }

    /// <summary>
    /// Показывает справку по использованию программы.
    /// </summary>
    private static void ShowUsage()
    {
        _logger.Info("ИСПОЛЬЗОВАНИЕ:");
        _logger.Info("  DocToScan.exe <файл1> [<файл2> ...]");
        _logger.Info("  DocToScan.exe /install    - установка в контекстное меню");
        _logger.Info("  DocToScan.exe /uninstall  - удаление из контекстного меню");
        _logger.Info("");
        _logger.Info("ПРИМЕРЫ:");
        _logger.Info("  DocToScan.exe \"C:\\Docs\\report.docx\"");
        _logger.Info("  DocToScan.exe \"file1.pdf\" \"file2.docx\"");
        _logger.Info("");
        _logger.Info("ПОДДЕРЖИВАЕМЫЕ ФОРМАТЫ:");
        _logger.Info("  - PDF (*.pdf)");
        _logger.Info("  - Word (*.docx, *.doc)");
        _logger.Info("");
        _logger.Info("КОНФИГУРАЦИЯ:");
        _logger.Info("  Файл config.xml в папке программы");
    }

    /// <summary>
    /// Ожидает нажатия клавиши перед выходом, если приложение запущено не из консоли.
    /// </summary>
    private static void WaitForExit()
    {
        // Проверяем, запущено ли приложение из консоли
        bool isConsoleAttached = Console.WindowHeight > 0;

        if (isConsoleAttached)
        {
            _logger.Info("");
            _logger.Info("Нажмите любую клавишу для выхода...");
            Console.ReadKey();
        }
    }
}