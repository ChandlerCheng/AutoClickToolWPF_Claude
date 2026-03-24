using System;
using System.Drawing;
using System.Windows;
using System.Windows.Input;
using System.IO;

namespace AutoClickTool_WPF.Tool
{ 
    public class DebugFunction
    {
        public static bool IsDebugMsg = false;
        public static bool IsDebugDownloadImg = false;
        public static string feedBackFunctionTitle(int functionIndex)
        {
            string result = string.Empty;
            switch (functionIndex)
            {
                case 2:
                    result = Application.Current.Resources["tabAutoBattle"].ToString();
                    break;
                case 3:
                    result = Application.Current.Resources["tabAutoDefend"].ToString();
                    break;
                case 4:
                    result = Application.Current.Resources["tabEnterBattleKey"].ToString();
                    break;
                case 5:
                    result = Application.Current.Resources["tabBuffTarget"].ToString(); 
                    break;
                case 6:
                    result = Application.Current.Resources["tabAuxiliaryFunctions"].ToString();
                    break;
                case 7:
                    result = Application.Current.Resources["tabSkillTraining"].ToString();
                    break;
                default:
                    result = "Error";
                    break;
            }
            return result;
        }
        public static string feedBackKeyString(Key inputKey)
        {
            string result = string.Empty;
            switch (inputKey)
            {
                case Key.F5:
                    result = "F5";
                    break;
                case Key.F6:
                    result = "F6";
                    break;
                case Key.F7:
                    result = "F7";
                    break;
                case Key.F8:
                    result = "F8";
                    break;
                case Key.F9:
                    result = "F9";
                    break;
                case Key.F10:
                    result = "F10";
                    break;
                case Key.F11:
                    result = "F11";
                    break;
                case Key.F12:
                    result = "F12";
                    break;
                default:
                    result = "";
                    break;
            }
            return result;
        }
        public static void captureAllEnemyDotScreen()
        {
            int xOffset = Coordinate.windowBoxLineOffset + Coordinate.windowTop[0];
            int yOffset = Coordinate.windowHOffset + 1 + Coordinate.windowTop[1];

            int x, y;

            for (int i = 0; i < 10; i++)
            {
                x = Coordinate.checkEnemy[i, 0] + xOffset;
                y = Coordinate.checkEnemy[i, 1] + yOffset;
                using (Bitmap bmp = BitmapFunction.CaptureScreen(x, y, 1, 1))
                {
                    bmp.Save("Enemy_" + i + "_" + "x" + x + "_" + "y" + y + "_" + ".bmp");
                }
            }
        }
        public static void captureTargetScreen(int in_x, int in_y, int width, int height)
        {
            // 取得怪物比較點時 y軸必須得加一
            int xOffset = Coordinate.windowBoxLineOffset + Coordinate.windowTop[0];
            int yOffset = Coordinate.windowHOffset + Coordinate.windowTop[1];
            int x = xOffset + in_x;
            int y = yOffset + in_y;
            using (Bitmap GetBmp = BitmapFunction.CaptureScreen(x, y, width, height))
            {
                GetBmp.Save("CaptureBitmap_" + "_" + "x" + in_x + "_" + "y" + in_y + "_" + ".bmp");
            }
        }
        public static void captureAllItemScreen()
        {
            int xo, yo;
            int xOffset = Coordinate.windowBoxLineOffset + Coordinate.windowTop[0];
            int yOffset = Coordinate.windowHOffset + Coordinate.windowTop[1];
            int i;
            int getMax = 16;

            // 取得物品欄檢查點  479,236  抓小塊 (物品欄)
            using (Bitmap ItemCheckBmp = BitmapFunction.CaptureScreen(479 + xOffset, 236 + yOffset, 20, 20))
            {
                ItemCheckBmp.Save("ItemCheck" + "_Bitmap_" + "_" + "x" + 479 + "_" + "y" + 236 + "_" + ".bmp");
            }
            // 取得所有物品欄位
            for (i = 0; i < getMax; i++)
            {
                GameFunction.GetItemCoor(i, out xo, out yo);
                using (Bitmap ItemGetBmp = BitmapFunction.CaptureScreen(xo + xOffset, yo + yOffset, 20, 20))
                {
                    ItemGetBmp.Save(i + "_Bitmap_" + "_" + "x" + xo + "_" + "y" + yo + "_" + ".bmp");
                }
            }
        }

        public static void captureAllNeedBmp()
        {
            int xOffset = Coordinate.windowBoxLineOffset + Coordinate.windowTop[0];
            int yOffset = Coordinate.windowHOffset + Coordinate.windowTop[1];
            int xNormalCheck = xOffset + 115;
            int yNormalCheck = yOffset + 1;

            int xBattlePlayer = xOffset + 766;
            int yBattlePlayer = yOffset + 98;

            int xItemList = xOffset + 479;
            int yItemList = yOffset + 236;

            MessageBox.Show($"開始取得視窗所有圖片模式 , 請在指示後進入對應的狀態再按下確認 , 期間請勿遮擋遊戲視窗");
            MessageBox.Show($"請處於非戰鬥狀態 , 以便擷取一般狀態檢查點");
            //NormalCheck
            using (Bitmap screenshot_NormalCheckPoint = BitmapFunction.CaptureScreen(xNormalCheck, yNormalCheck, 9, 30))
            {
                screenshot_NormalCheckPoint.Save("win10_NormalCheckPoint_800x600" + ".bmp");
            }

            MessageBox.Show($"請處於戰鬥狀態的人物視角 , 以便擷取人物戰鬥狀態檢查點");
            // BattleCheck_Player
            using (Bitmap screenshot_keyBarPlayer = BitmapFunction.CaptureScreen(xBattlePlayer, yBattlePlayer, 33, 34))
            {
                screenshot_keyBarPlayer.Save("win10_fighting_keybar_player_800x600" + ".bmp");
            }

            MessageBox.Show($"請處於戰鬥狀態的寵物視角 , 以便擷取寵物戰鬥狀態檢查點");
            // BattleCheck_Pet
            using (Bitmap screenshot_keyBarPet = BitmapFunction.CaptureScreen(xBattlePlayer, yBattlePlayer, 33, 34))
            {
                screenshot_keyBarPet.Save("win10_fighting_keybar_pet_800x600" + ".bmp");
            }

            MessageBox.Show($"請處於非戰鬥狀態並打開物品欄 , 以便擷取物品欄檢查點");
            // ItemTimeCheck
            using (Bitmap screenshot_itemTimeBitmap = BitmapFunction.CaptureScreen(xItemList, yItemList, 20, 20))
            {
                screenshot_itemTimeBitmap.Save("win10_ItemCheckPoint_800x600" + ".bmp");
            }

            //MessageBox.Show($"請處於非戰鬥狀態並打開物品欄 , 將魔晶寶箱置於第三排第一格 ,以便擷取物品檢查點");
            // CheckItemCB_Coor

            //MessageBox.Show($"請處於非戰鬥狀態並打開物品欄 , 將調整藥丸置於第三排第一格 ,以便擷取物品檢查點");
            //CheckItemAP_Coor

            //MessageBox.Show($"請使用調整藥丸 , 並將畫面停留在使用藥物詢問視窗");
            //CheckItemAP_Yes


            //MessageBox.Show($"");
            //MessageBox.Show($"");
        }
    }
    public static class Log
    {
        private static readonly object _lock = new object();
        private static string _logFilePath;
        public static void Init(string exeBaseDir = null)
        {
            string exePath = (exeBaseDir ?? AppDomain.CurrentDomain.BaseDirectory).TrimEnd('\\');
            string timestamp = DateTime.Now.ToString("yyyyMMdd"); // 以天為單位
            string fileName = $"AutoClickToolWPF_{timestamp}.log";
            string folder = Path.Combine(exePath, "Logs");
            if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
            _logFilePath = Path.Combine(folder, fileName);

            // 可選：在 ProcessExit 時做最後 flush（這個版本每次都關檔，實際上不需要）
            AppDomain.CurrentDomain.ProcessExit += (s, e) => { /* no-op */ };
        }
        public static void Append(string message)
        {
            if (string.IsNullOrEmpty(_logFilePath)) return;
            if (DebugFunction.IsDebugMsg != true) return;
            string timeStamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string line = $"[{timeStamp}] {message}";

            // 以 lock 確保多執行緒一次只會有一個 Writer 開檔寫入
            lock (_lock)
            {
                // 用 FileStream 控制 FileShare，避免其他程式讀檔被鎖住
                using (var fs = new FileStream(_logFilePath, FileMode.Append, FileAccess.Write, FileShare.Read))
                using (var writer = new StreamWriter(fs, new System.Text.UTF8Encoding(true)))
                {
                    writer.WriteLine(line);
                    writer.Flush();
                    // using 離開即關檔
                }
            }
        }
        public static string FormatElapsedTime(long totalMilliseconds)
        {
            if (totalMilliseconds < 1000)
            {
                double seconds = totalMilliseconds / 1000.0;
                return seconds.ToString("0.0") + "秒";
            }
            else
            {
                TimeSpan ts = TimeSpan.FromMilliseconds(totalMilliseconds);
                return string.Format("{0:00}:{1:00}:{2:00}",
                                     (int)ts.TotalHours,
                                     ts.Minutes,
                                     ts.Seconds);
            }
        }
    }
}
