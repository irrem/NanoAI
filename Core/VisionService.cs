using System;
using System.Drawing;
using System.Windows.Forms;

namespace NanAI.Core
{
    public static class VisionService
    {
        public static void CaptureAndSaveScreen()
        {
            try
            {
                Rectangle bounds = Screen.PrimaryScreen.Bounds;
                using Bitmap bitmap = new Bitmap(bounds.Width, bounds.Height);
                using Graphics g = Graphics.FromImage(bitmap);
                g.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);
                string filePath = $"screenshot_{DateTime.Now:yyyyMMdd_HHmmss}.png";
                bitmap.Save(filePath);
                Console.WriteLine($"Screenshot saved: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to capture screenshot: {ex.Message}");
            }
        }
    }
}
