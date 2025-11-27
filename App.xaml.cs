using Application = Microsoft.Maui.Controls.Application;

namespace InvoiceProcessor
{
    public partial class App : Application
    {
        public App()
        {
            InitializeComponent();
            MainPage = new AppShell();
        }

        protected override Window CreateWindow(IActivationState? activationState)
        {
            var window = base.CreateWindow(activationState);

            // === 设置默认窗口大小 ===
            const int newWidth = 900;  // 宽度
            const int newHeight = 800; // 高度

            window.Width = newWidth;
            window.Height = newHeight;

            // 可选：设置最小尺寸，防止用户拖得太小看不见
            window.MinimumWidth = 800;
            window.MinimumHeight = 600;

            window.X = (DeviceDisplay.MainDisplayInfo.Width / DeviceDisplay.MainDisplayInfo.Density - newWidth) / 2;
            window.Y = (DeviceDisplay.MainDisplayInfo.Height / DeviceDisplay.MainDisplayInfo.Density - newHeight) / 2;

            return window;
        }
    }
}