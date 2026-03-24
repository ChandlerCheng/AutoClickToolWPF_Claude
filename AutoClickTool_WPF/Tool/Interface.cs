using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Threading;

namespace AutoClickTool_WPF.Tool
{
    public class KeyboardSimulator
    {
        // 匯入 user32.dll 中的 keybd_event 函數
        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        // 定義按鍵事件的標誌
        private const int KEYEVENTF_EXTENDEDKEY = 0x0001;
        private const int KEYEVENTF_KEYUP = 0x0002;

        // 模擬按下按鍵
        public static void KeyDown(Key key)
        {
            byte keyCode = (byte)KeyInterop.VirtualKeyFromKey(key);
            Application.Current.Dispatcher.Invoke(() =>
            {
                keybd_event(keyCode, 0, KEYEVENTF_EXTENDEDKEY, UIntPtr.Zero);
            });
        }

        // 模擬釋放按鍵
        public static void KeyUp(Key key)
        {
            byte keyCode = (byte)KeyInterop.VirtualKeyFromKey(key);
            Application.Current.Dispatcher.Invoke(() =>
            {
                keybd_event(keyCode, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, UIntPtr.Zero);
            });
        }

        // 模擬按下並釋放按鍵
        public static void KeyPress(Key key)
        {
            byte keyCode = (byte)KeyInterop.VirtualKeyFromKey(key);
            Application.Current.Dispatcher.Invoke(() =>
            {
                KeyDown(key);
                Thread.Sleep(GameScript.keyPressDelay); // 延遲一段時間
                KeyUp(key);
            });
        }
    }
    public class MouseSimulator
    {
        public struct POINT
        {
            public int X;
            public int Y;
        }
        // 匯入 GetCursorPos 函數
        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out POINT lpPoint);
        // 匯入User32.dll中的函數
        [DllImport("user32.dll")]
        public static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetCursorPos(int X, int Y);

        // 模擬滑鼠左鍵按下及釋放
        private const int MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const int MOUSEEVENTF_LEFTUP = 0x0004;
        // 模擬滑鼠右鍵按下及釋放
        private const int MOUSEEVENTF_RIGHTDOWN = 0x0008;
        private const int MOUSEEVENTF_RIGHTUP = 0x0010;
        private const int MOUSEEVENTF_ABSOLUTE = 0x8000;
        // WPF UI 執行緒安全的滑鼠左鍵按下方法
        public static void LeftMousePress(int x, int y)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                // 移動滑鼠到指定座標
                SetCursorPos(x, y);
                // 模擬滑鼠左鍵按下
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_ABSOLUTE, 0, 0, 0, UIntPtr.Zero);
                Thread.Sleep(GameScript.keyPressDelayLong); // 延遲一段時間
                mouse_event(MOUSEEVENTF_LEFTUP | MOUSEEVENTF_ABSOLUTE, 0, 0, 0, UIntPtr.Zero);
            });
        }
        // WPF UI 執行緒安全的滑鼠右鍵按下方法
        public static void RightMousePress(int x, int y)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                // 移動滑鼠到指定座標
                SetCursorPos(x, y);

                // 模擬滑鼠右鍵按下
                mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, UIntPtr.Zero);
                Thread.Sleep(GameScript.keyPressDelayLong); // 延遲一段時間

                // 模擬滑鼠右鍵釋放
                mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, UIntPtr.Zero);
            });
        }
        // WPF UI 執行緒安全的移動滑鼠位置
        public static void MoveMouseTo(int x, int y)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                // 移動滑鼠到指定座標
                SetCursorPos(x, y);
            });
        }

        public static void GetCurrentXY(out int x, out int y)
        {
            // 取得當前滑鼠位置
            if (GetCursorPos(out POINT pt))
            {
                x = pt.X;
                y = pt.Y;
            }
            else
            {
                x = 0;
                y = 0;
            }
            // 輸出目前座標到控制台
            Console.WriteLine("當前滑鼠位置: X = {0}, Y = {1}", x, y);
        }
    }
}
