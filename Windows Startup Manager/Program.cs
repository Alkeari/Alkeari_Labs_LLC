using Avalonia;
using System;

namespace Windows_Startup_Manager
{
    internal static class Program
    {
        [STAThread]
        public static void Main(string[] args) => BuildAvaloniaApp()
            .StartWithClassicDesktopLifetime(args);

        private static AppBuilder BuildAvaloniaApp() => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
    }
}