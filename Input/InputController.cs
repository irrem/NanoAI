using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using WindowsInput.Native;
using WindowsInput;

namespace NanAI.Input
{
    public static class InputController
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private const uint WM_CLOSE = 0x0010;

        public static void CloseActiveWindow()
        {
            IntPtr hWnd = GetForegroundWindow();
            if (hWnd != IntPtr.Zero)
            {
                PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                Console.WriteLine("Active window closed.");
            }
            else
            {
                Console.WriteLine("No active window found.");
            }
        }

        public static void SimulateKeyPress(VirtualKeyCode keyCode)
        {
            var sim = new InputSimulator();
            sim.Keyboard.KeyPress(keyCode);
        }
    }
}
