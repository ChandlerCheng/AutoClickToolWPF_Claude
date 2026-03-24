using System;
using System.Runtime.InteropServices;
using System.Windows;

using AutoClickTool_WPF.Tool;

namespace AutoClickTool_WPF
{
    public class SystemSetting
    {
        // Windows 版本判斷
        public static bool isWin10 = false;
        public static bool isDebug = false;
        // 引入 RegisterHotKey API
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        // 引入 UnregisterHotKey API
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        // 定義熱鍵 ID，確保唯一
        public const int HOTKEY_SCRIPT_EN = 9000;

        // 匯入 FindWindow
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        // 匯入 GetWindowRect
        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        // 定義 RECT 結構體，對應於 Windows API 的 RECT
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public static bool GetWindowCoordinates(string windowTitle)
        {
            IntPtr windowHandle = FindWindow(null, windowTitle);
            if (windowHandle != IntPtr.Zero)
            {
                if (GetWindowRect(windowHandle, out RECT GameWindowsInfor))
                {
                    /*
                        遊戲本體800x600
                        視窗總長816x638
                        視窗邊框約8
                     */
                    Coordinate.windowTop[0] = GameWindowsInfor.Left;
                    Coordinate.windowTop[1] = GameWindowsInfor.Top;
                    Coordinate.windowBottom[0] = GameWindowsInfor.Right;
                    Coordinate.windowBottom[1] = GameWindowsInfor.Bottom;
                    Coordinate.windowHeigh = GameWindowsInfor.Bottom - GameWindowsInfor.Top;
                    Coordinate.windowWidth = GameWindowsInfor.Right - GameWindowsInfor.Left;
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }
        public static void GetGameWindow()
        {
            // 抓取視窗位置 & 視窗長寬數值
            if (!GetWindowCoordinates("FairyLand"))
            {
                Coordinate.IsGetWindows = false;
                return;
            }
            Coordinate.IsGetWindows = true;
            Coordinate.CalculateWalkPoint();
            /*
              中心怪物位置會位於 800x600中的 275,290的位置
              遊戲本體800x600
              視窗總長 Windows 7 = 816x638  , Windows 10 = 816x639
              視窗邊框約8
              視窗標題列約30

               友軍前排3號位529,387
               敵軍前排3號位275,290
            */
            int x_enemy3, y_enemy3;
            int xOffset_enemy3 = Coordinate.windowBoxLineOffset + 275;
            int yOffset_enemy3 = Coordinate.windowHOffset + 290;

            x_enemy3 = Coordinate.windowTop[0] + xOffset_enemy3;
            y_enemy3 = Coordinate.windowTop[1] + yOffset_enemy3;

            Coordinate.CalculateAllEnemy(x_enemy3, y_enemy3);

            int x_friends3, y_friends3;
            int xOffset_friends3 = Coordinate.windowBoxLineOffset + 529;
            int yOffset_friends3 = Coordinate.windowHOffset + 387;

            x_friends3 = Coordinate.windowTop[0] + xOffset_friends3;
            y_friends3 = Coordinate.windowTop[1] + yOffset_friends3;

            Coordinate.CalculateAllFriends(x_friends3, y_friends3);

        }

        // 加載對應的語系資源字典
        public static void LoadResourceDictionary(string culture)
        {
            // 移除舊的資源字典
            ResourceDictionary oldDict = null;
            foreach (ResourceDictionary dict in Application.Current.Resources.MergedDictionaries)
            {
                if (dict.Source != null && (dict.Source.OriginalString.Contains("en-US.xaml") || dict.Source.OriginalString.Contains("zh-TW.xaml")))
                {
                    oldDict = dict;
                    break;
                }
            }

            if (oldDict != null)
            {
                Application.Current.Resources.MergedDictionaries.Remove(oldDict);
            }

            // 加載新的資源字典
            var newDict = new ResourceDictionary();
            newDict.Source = new Uri($"/AutoClickTool_WPF;component/Language/{culture}", UriKind.Relative);
            Application.Current.Resources.MergedDictionaries.Add(newDict);
        }
    }
}
