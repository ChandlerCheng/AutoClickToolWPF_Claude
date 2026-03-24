using System;
using System.Drawing;
using System.Threading;
using System.Windows.Input;

using AutoClickTool_WPF.Tool;

namespace AutoClickTool_WPF
{
    public class GameFunction
    {
        public static void GameRightMousePress(int x, int y)
        {
            MouseSimulator.RightMousePress(x + Coordinate.windowBoxLineOffset, y + Coordinate.windowHOffset);
        }
        public static void GameLeftMousePress(int x, int y)
        {
            MouseSimulator.LeftMousePress(x + Coordinate.windowBoxLineOffset, y + Coordinate.windowHOffset);
        }
        public static void GameMoveMouseTo(int x, int y)
        {
            MouseSimulator.MoveMouseTo(x + Coordinate.windowBoxLineOffset, y + Coordinate.windowHOffset);
        }
        public static void AttackOnTarget(int x, int y)
        {
            /*
                滑鼠移動到指定座標後 , 按下指定熱鍵 , 點下左鍵
             */
            MouseSimulator.LeftMousePress(x, y);
            Thread.Sleep(GameScript.keyPressDelayLong);
        }
        public static void castSpellOnTarget(int x, int y, Key keyCode, int delay)
        {
            /*
                滑鼠移動到指定座標後 , 按下指定熱鍵 , 點下左鍵
             */
            KeyboardSimulator.KeyPress(keyCode);
            Thread.Sleep(GameScript.keyPressDelayLong);
            if (GameScript.isLikeHuman)
            {
                int randomMs = 0;
                Random random = new Random();
                randomMs = random.Next(4, 5) * 100;
                Thread.Sleep(randomMs);
            }
            MouseSimulator.LeftMousePress(x, y);
            Thread.Sleep(delay);
        }
        public static void openToralItemScene()
        {
            // 按下 Ctrl 鍵
            KeyboardSimulator.KeyDown(Key.LeftCtrl);
            Thread.Sleep(GameScript.keyPressDelay);
            // 按下 B 鍵
            KeyboardSimulator.KeyDown(Key.B);
            Thread.Sleep(GameScript.keyPressDelay);
            // 釋放 B 鍵
            KeyboardSimulator.KeyUp(Key.B);
            Thread.Sleep(GameScript.keyPressDelay);
            // 釋放 Ctrl 鍵
            KeyboardSimulator.KeyUp(Key.LeftCtrl);
            Thread.Sleep(GameScript.keyPressDelay);
        }
        public static bool NormalCheck()
        {
            int x_LU, y_LU;
            int xOffset_LU = Coordinate.windowBoxLineOffset + 115;
            int yOffset_LU = Coordinate.windowHOffset + 1;
            x_LU = Coordinate.windowTop[0] + xOffset_LU;
            y_LU = Coordinate.windowTop[1] + yOffset_LU;
            Log.Append($"=============================");
            Log.Append($"座標 X = ['{x_LU}'] Y=['{y_LU}']");
            Bitmap NormalCheckPoint;

            if (SystemSetting.isWin10 == true)
                NormalCheckPoint = Properties.Resources.win10_NormalCheckPoint_800x600;
            else
                NormalCheckPoint = Properties.Resources.win7_NormalCheckPoint;

            GameScript.mousePostionDefault();
            // 從畫面上擷取指定區域的圖像
            using (Bitmap screenshot_NormalCheckPoint = BitmapFunction.CaptureScreen(x_LU, y_LU, 9, 30))
            {
                // 比對圖像
                double Final_NCP = BitmapFunction.CompareImages(screenshot_NormalCheckPoint, NormalCheckPoint);
                Log.Append($"一般狀態檢測值 '{Final_NCP}')\n");
                Log.Append($"=============================");

                if (Final_NCP > 80)
                    return true;
                else
                    return false;
            }
        }
        public static bool BattleCheck_Player()
        {
            int x_key, y_key;
            int xOffset_key = Coordinate.windowBoxLineOffset + 766;
            int yOffset_key = Coordinate.windowHOffset + 98;
            Bitmap fight_keybarBMP;

            x_key = Coordinate.windowTop[0] + xOffset_key;
            y_key = Coordinate.windowTop[1] + yOffset_key;
            Log.Append($"=============================");
            Log.Append($"座標 X = ['{x_key}'] Y=['{y_key}']");

            if (SystemSetting.isWin10 == true)
            {
                fight_keybarBMP = Properties.Resources.win10_fighting_keybar_player_800x600;
            }
            else
                fight_keybarBMP = Properties.Resources.win7_fighting_keybar_player;
            GameScript.mousePostionDefault();
            // 從畫面上擷取指定區域的圖像
            using (Bitmap screenshot_keyBar = BitmapFunction.CaptureScreen(x_key, y_key, 33, 34))
            {
                // 比對圖像
                double Final_KeyBar = BitmapFunction.CompareImages(screenshot_keyBar, fight_keybarBMP);
                Log.Append($"玩家檢測值為 '{Final_KeyBar}')\n");
                Log.Append($"=============================");

                if (Final_KeyBar > 80)
                    return true;
                else
                    return false;
            }
        }
        public static bool BattleCheck_Pet()
        {
            int x_key, y_key;
            int xOffset_key = Coordinate.windowBoxLineOffset + 766;
            int yOffset_key = Coordinate.windowHOffset + 98;
            Bitmap fight_keybarPetBMP;

            x_key = Coordinate.windowTop[0] + xOffset_key;
            y_key = Coordinate.windowTop[1] + yOffset_key;
            Log.Append($"=============================");
            Log.Append($"座標 X = ['{x_key}'] Y=['{y_key}']");

            GameScript.mousePostionDefault();

            if (SystemSetting.isWin10 == true)
            {
                fight_keybarPetBMP = Properties.Resources.win10_fighting_keybar_pet_800x600;
            }
            else
                fight_keybarPetBMP = Properties.Resources.win7_fighting_keybar_pet;
            // 從畫面上擷取指定區域的圖像
            using (Bitmap screenshot_keyBarPet = BitmapFunction.CaptureScreen(x_key, y_key, 33, 34))
            {
                // 比對圖像
                double Final_KeyBar = BitmapFunction.CompareImages(screenshot_keyBarPet, fight_keybarPetBMP);
                Log.Append($"寵物檢測值為 '{Final_KeyBar}')\n");
                Log.Append($"=============================");

                if (Final_KeyBar > 80)
                    return true;
                else
                    return false;
            }
        }
        public static void pressDefendButton()
        {
            /*
                有小bug , 會變成點擊原先停點上的怪物 , 而不是指向防禦按鈕

                20240510 : 加入 Cursor.Position = new System.Drawing.Point(x, y); 才確保會移動到正確位置上。
             */
            int xOffset = Coordinate.windowBoxLineOffset + Coordinate.windowTop[0];
            int yOffset = Coordinate.windowHOffset + 1 + Coordinate.windowTop[1];
            int x, y;
            x = 779 + xOffset;
            y = 68 + yOffset;
            MouseSimulator.LeftMousePress(x, y);
            Thread.Sleep(GameScript.keyPressDelay);
            GameScript.mousePostionDefault();
        }
        public static void pressRecheckBag()
        {
            /*
                重新整理背包
             */
            int xOffset = Coordinate.windowBoxLineOffset + Coordinate.windowTop[0];
            int yOffset = Coordinate.windowHOffset + 1 + Coordinate.windowTop[1];
            int x, y;
            x = 776 + xOffset;
            y = 510 + yOffset;
            MouseSimulator.LeftMousePress(x, y);
            Thread.Sleep(GameScript.keyPressDelay);
            GameScript.mousePostionDefault();
        }
        public static int getEnemyCoor()
        {
            int result = 0;
            if (BattleCheck_Player() == true || BattleCheck_Pet() == true)
            {
                int xOffset = Coordinate.windowBoxLineOffset + Coordinate.windowTop[0];
                int yOffset = Coordinate.windowHOffset + 1 + Coordinate.windowTop[1];

                Bitmap[] enemyGetBmp = new Bitmap[10];
                Bitmap[] enemyGetBmp_2 = new Bitmap[10];
                System.Drawing.Color EnemyExistColor;

                try
                {
                    for (int i = 0; i < 10; i++)
                        enemyGetBmp[i] = BitmapFunction.CaptureScreen(Coordinate.checkEnemy[i, 0] + xOffset, Coordinate.checkEnemy[i, 1] + yOffset, 1, 1);
                    for (int i = 0; i < 10; i++)
                        enemyGetBmp_2[i] = BitmapFunction.CaptureScreen(Coordinate.checkEnemy_2[i, 0] + xOffset, Coordinate.checkEnemy_2[i, 1] + yOffset, 1, 1);

                    // 184 188 184
                    if (SystemSetting.isWin10 == true)
                        EnemyExistColor = System.Drawing.Color.FromArgb(189, 190, 189);
                    else
                        EnemyExistColor = System.Drawing.Color.FromArgb(255, 255, 255);

                    for (int i = 0; i < 10; i++)
                    {
                        // 一次檢查 , 已知史萊姆系會無法判斷
                        double EnemyExistRatio = BitmapFunction.CalculateColorRatio(enemyGetBmp[i], EnemyExistColor);
                        if (EnemyExistRatio > 0)
                        {
                            result = i + 1;
                            return result;
                        }
                        else
                        {
                            Log.Append($"1_檢查第'{i}'位置時(X={Coordinate.checkEnemy[i, 0]},Y={Coordinate.checkEnemy[i, 1]}) , 比對值為'{EnemyExistRatio}'");
                            // 二次檢查 , 已知史萊姆在對話欄有字時 , 第一位會失效
                            if (SystemSetting.isWin10 == true)
                                EnemyExistColor = System.Drawing.Color.FromArgb(214, 211, 214);
                            else
                                EnemyExistColor = System.Drawing.Color.FromArgb(255, 255, 255);

                            double EnemyExistRatio_2 = BitmapFunction.CalculateColorRatio(enemyGetBmp_2[i], EnemyExistColor);
                            if (EnemyExistRatio_2 > 0)
                            {
                                result = i + 1;
                                return result;
                            }
                            Log.Append($"2_檢查第'{i}'位置時(X={Coordinate.checkEnemy_2[i, 0]},Y={Coordinate.checkEnemy_2[i, 1]}) , 比對值為'{EnemyExistRatio}'");
                        }

                    }
                    Log.Append($"無法取得怪物序列");

                    return 0;
                }
                finally
                {
                    foreach (var bmp in enemyGetBmp) bmp?.Dispose();
                    foreach (var bmp in enemyGetBmp_2) bmp?.Dispose();
                }
            }

            if (DebugFunction.IsDebugMsg == true)
                Log.Append($"非戰鬥狀態");
            return 0;
        }
        public static bool ItemTimeCheck()
        {

            int x_ItemCP, y_ItemCP;
            int xOffset_LU = Coordinate.windowBoxLineOffset + 479;
            int yOffset_LU = Coordinate.windowHOffset + 236;
            x_ItemCP = Coordinate.windowTop[0] + xOffset_LU;
            y_ItemCP = Coordinate.windowTop[1] + yOffset_LU;
            Bitmap itemTimeBitmap;

            itemTimeBitmap = Properties.Resources.Win7_ItemCheckPoint_479_236_20x20;

            // 從畫面上擷取指定區域的圖像
            using (Bitmap screenshot_itemTimeBitmap = BitmapFunction.CaptureScreen(x_ItemCP, y_ItemCP, 20, 20))
            {
                // 比對圖像
                double Final_ICP = BitmapFunction.CompareImages(screenshot_itemTimeBitmap, itemTimeBitmap);

                Log.Append($"物品欄檢測值為 '{Final_ICP}')\n");

                if (Final_ICP > 50)
                    return true;
                else
                    return false;
            }
        }
        public static void GetItemCoor(int row, int column, out int x, out int y)
        {
            // 初始化初始的 x 和 y
            int x_first = 515;
            int y_first = 320;
            int x_rowOffset = 0;
            int y_rowOffset = 0;

            // 每列的水平偏移
            int x_RowOffset = 57;
            // 每行的垂直偏移
            int y_columnOffset = 51;

            // 計算 y 值（根據 row 計算 y_rowOffset）
            y_rowOffset = (column - 1) * y_columnOffset;
            y = y_first + y_rowOffset;

            // 計算 x 值（根據 column 計算 x_rowOffset）
            x_rowOffset = (row - 1) * x_RowOffset;
            x = x_first + x_rowOffset;
        }
        public static void GetItemCoor(int i, out int x, out int y)
        {
            // 初始化初始的 x 和 y
            int x_first = 515;
            int y_first = 320;
            int x_rowOffset = 0;
            int y_rowOffset = 0;

            // 每列的水平偏移
            int x_RowOffset = 57;
            // 每行的垂直偏移
            int y_columnOffset = 51;

            // 假設是每列4個項目，根據i來計算row和column
            int itemsPerRow = 4;

            // 計算 row 和 column
            int row = (i % itemsPerRow) + 1;     // 行對應的索引 (row)
            int column = (i / itemsPerRow) + 1;  // 列對應的索引 (column)

            // 計算 y 值（根據 column 計算 y_rowOffset）
            y_rowOffset = (column - 1) * y_columnOffset;
            y = y_first + y_rowOffset;

            // 計算 x 值（根據 row 計算 x_rowOffset）
            x_rowOffset = (row - 1) * x_RowOffset;
            x = x_first + x_rowOffset;
        }
        public static void GetTradeItemCoor(int i, out int x, out int y)
        {
            // 初始化初始的 x 和 y
            int x_first = 441;
            int y_first = 199;
            int x_rowOffset = 0;
            int y_rowOffset = 0;

            // 每列的水平偏移
            int x_RowOffset = 54;
            // 每行的垂直偏移
            int y_columnOffset = 58;

            // 假設是每列4個項目，根據i來計算row和column
            int itemsPerRow = 5;

            // 計算 row 和 column
            int row = (i % itemsPerRow) + 1;     // 行對應的索引 (row)
            int column = (i / itemsPerRow) + 1;  // 列對應的索引 (column)

            // 計算 y 值（根據 column 計算 y_rowOffset）
            if (column == 1)
                y_rowOffset = 0;
            else if (column == 2)
                y_rowOffset = 58;
            else if (column == 3)
                y_rowOffset = 111;
            else if (column == 4)
                y_rowOffset = 165;

            y = y_first + y_rowOffset;

            // 計算 x 值（根據 row 計算 x_rowOffset）
            x_rowOffset = (row - 1) * x_RowOffset;
            x = x_first + x_rowOffset;
        }
        public static void GetTradeItemSelectCoor(int i, out int x, out int y)
        {
            // 初始化初始的 x 和 y
            int x_first = 425;
            int y_first = 184;
            int x_rowOffset = 0;
            int y_rowOffset = 0;

            // 每列的水平偏移
            int x_RowOffset = 54;
            // 每行的垂直偏移
            int y_columnOffset = 58;

            // 假設是每列4個項目，根據i來計算row和column
            int itemsPerRow = 5;

            // 計算 row 和 column
            int row = (i % itemsPerRow) + 1;     // 行對應的索引 (row)
            int column = (i / itemsPerRow) + 1;  // 列對應的索引 (column)

            // 計算 y 值（根據 column 計算 y_rowOffset）
            if (column == 1)
                y_rowOffset = 0;
            else if (column == 2)
                y_rowOffset = 58;
            else if (column == 3)
                y_rowOffset = 111;
            else if (column == 4)
                y_rowOffset = 165;

            y = y_first + y_rowOffset;

            // 計算 x 值（根據 row 計算 x_rowOffset）
            x_rowOffset = (row - 1) * x_RowOffset;
            x = x_first + x_rowOffset;
        }

        public static void GetInventoryItemCoor(int i, out int x, out int y)
        {
            // 初始化初始的 x 和 y
            int x_first = 528;
            int y_first = 337;
            int x_rowOffset = 0;
            int y_rowOffset = 0;

            // 每列的水平偏移
            int x_RowOffset = 57;
            // 每行的垂直偏移
            int y_columnOffset = 51;

            // 假設是每列4個項目，根據i來計算row和column
            int itemsPerRow = 4;

            // 計算 row 和 column
            int row = (i % itemsPerRow) + 1;     // 行對應的索引 (row)
            int column = (i / itemsPerRow) + 1;  // 列對應的索引 (column)


            // 計算 y 值（根據 column 計算 y_rowOffset）
            y_rowOffset = (column - 1) * y_columnOffset;
            y = y_first + y_rowOffset;

            // 計算 x 值（根據 row 計算 x_rowOffset）
            x_rowOffset = (row - 1) * x_RowOffset;
            x = x_first + x_rowOffset;
        }
        public static void GetInventoryItemSelectedCoor(int i, out int x, out int y)
        {
            // 初始化初始的 x 和 y
            int x_first = 555;
            int y_first = 331;
            int x_rowOffset = 0;
            int y_rowOffset = 0;

            // 每列的水平偏移
            int x_RowOffset = 57;
            // 每行的垂直偏移
            int y_columnOffset = 51;

            // 假設是每列4個項目，根據i來計算row和column
            int itemsPerRow = 4;

            // 計算 row 和 column
            int row = (i % itemsPerRow) + 1;     // 行對應的索引 (row)
            int column = (i / itemsPerRow) + 1;  // 列對應的索引 (column)

            // 計算 y 值（根據 column 計算 y_rowOffset）
            y_rowOffset = (column - 1) * y_columnOffset;
            y = y_first + y_rowOffset;

            // 計算 x 值（根據 row 計算 x_rowOffset）
            x_rowOffset = (row - 1) * x_RowOffset;
            x = x_first + x_rowOffset;
        }
        public static void GetBankItemCoor(int i, out int x, out int y)
        {
            // 初始化初始的 x 和 y
            int x_first = 221;
            int y_first = 96;
            int x_rowOffset = 0;
            int y_rowOffset = 0;

            // 每列的水平偏移
            int x_RowOffset = 45;
            // 每行的垂直偏移
            int y_columnOffset = 44;

            // 假設是每列4個項目，根據i來計算row和column
            int itemsPerRow = 5;

            // 計算 row 和 column
            int row = (i % itemsPerRow) + 1;     // 行對應的索引 (row)
            int column = (i / itemsPerRow) + 1;  // 列對應的索引 (column)

            // 計算 y 值（根據 column 計算 y_rowOffset）
            y_rowOffset = (column - 1) * y_columnOffset;
            y = y_first + y_rowOffset;

            // 計算 x 值（根據 row 計算 x_rowOffset）
            x_rowOffset = (row - 1) * x_RowOffset;
            x = x_first + x_rowOffset;
        }
        public static void GetBankItemSelectedCoor(int i, out int x, out int y)
        {
            // 初始化初始的 x 和 y
            int x_first = 248;
            int y_first = 87;
            int x_rowOffset = 0;
            int y_rowOffset = 0;

            // 每列的水平偏移
            int x_RowOffset = 45;
            // 每行的垂直偏移
            int y_columnOffset = 44;

            // 假設是每列4個項目，根據i來計算row和column
            int itemsPerRow = 5;

            // 計算 row 和 column
            int row = (i % itemsPerRow) + 1;     // 行對應的索引 (row)
            int column = (i / itemsPerRow) + 1;  // 列對應的索引 (column)

            // 計算 y 值（根據 column 計算 y_rowOffset）
            y_rowOffset = (column - 1) * y_columnOffset;
            y = y_first + y_rowOffset;

            // 計算 x 值（根據 row 計算 x_rowOffset）
            x_rowOffset = (row - 1) * x_RowOffset;
            x = x_first + x_rowOffset;
        }

        public static bool CheckItemCB_Coor(int input, out int item_x, out int item_y)
        {
            int xo, yo;
            int xOffset = Coordinate.windowBoxLineOffset + 15;// 童話改版後物品欄已加入亂數像素點
            int yOffset = Coordinate.windowHOffset + 12;// 童話改版後物品欄已加入亂數像素點
            item_x = 0;
            item_y = 0;
            System.Drawing.Color PillExistColor = System.Drawing.Color.FromArgb(116, 85, 136);
            // 取得所有物品欄位
            GameFunction.GetItemCoor(input, out xo, out yo);

            double Final_ICP;
            using (Bitmap screenshot_ItemCBBmpGet = BitmapFunction.CaptureScreen(xo + xOffset, yo + yOffset, 1, 1))
            {
                // 比對圖像
                Final_ICP = BitmapFunction.CalculateColorRatio(screenshot_ItemCBBmpGet, PillExistColor);
            }

            Log.Append($"物品欄檢測值為 '{Final_ICP}')\n");

            if (Final_ICP > 0)
            {
                item_x = xo + xOffset;
                item_y = yo + yOffset;
                return true;
            }
            else
                return false;
        }
        public static bool CheckItemAP_Coor(int input, out int item_x, out int item_y)
        {
            int xo, yo;
            int xOffset = Coordinate.windowBoxLineOffset + 15;// 童話改版後物品欄已加入亂數像素點
            int yOffset = Coordinate.windowHOffset + 12;// 童話改版後物品欄已加入亂數像素點
            item_x = 0;
            item_y = 0;
            System.Drawing.Color PillExistColor = System.Drawing.Color.FromArgb(113, 148, 156);
            // 取得所有物品欄位
            GameFunction.GetItemCoor(input, out xo, out yo);

            double Final_ICP;
            using (Bitmap screenshot_ItemCBBmpGet = BitmapFunction.CaptureScreen(xo + xOffset, yo + yOffset, 1, 1))
            {
                // 比對圖像
                Final_ICP = BitmapFunction.CalculateColorRatio(screenshot_ItemCBBmpGet, PillExistColor);
            }

            Log.Append($"物品欄檢測值為 '{Final_ICP}')\n");

            if (Final_ICP > 0)
            {
                item_x = xo + xOffset;
                item_y = yo + yOffset;
                return true;
            }
            else
                return false;
        }
        public static bool CheckItemAP_Yes()
        {
            int x_key, y_key;
            int xOffset_key = Coordinate.windowBoxLineOffset + 249;
            int yOffset_key = Coordinate.windowHOffset + 282;
            Bitmap AP_UsingOK_BMP;

            x_key = Coordinate.windowTop[0] + xOffset_key;
            y_key = Coordinate.windowTop[1] + yOffset_key;

            AP_UsingOK_BMP = Properties.Resources.Win7_AdjustmentPill_Yes;
            // 從畫面上擷取指定區域的圖像
            using (Bitmap screenshot_AP_UsingOK = BitmapFunction.CaptureScreen(x_key, y_key, 10, 10))
            {
                // 比對圖像
                double Final_KeyBar = BitmapFunction.CompareImages(screenshot_AP_UsingOK, AP_UsingOK_BMP);

                Log.Append($"CheckItemAP_Yes 檢測值為 '{Final_KeyBar}')\n");

                if (Final_KeyBar > 80)
                    return true;
                else
                    return false;
            }
        }
        public static double GetBattleMpRatio()
        {
            double result = 0;
            int x, y;
            int xOffset_key = Coordinate.windowBoxLineOffset + 24;
            int yOffset_key = Coordinate.windowHOffset + 42;
            System.Drawing.Color MpExistColor;
            x = Coordinate.windowTop[0] + xOffset_key;
            y = Coordinate.windowTop[1] + yOffset_key;
            MpExistColor = System.Drawing.Color.FromArgb(73, 206, 254);
            using (Bitmap screenshot_Mp = BitmapFunction.CaptureScreen(x, y, 90, 1))
            {
                result = BitmapFunction.CalculateColorRatio(screenshot_Mp, MpExistColor);
            }
            return result;
        }
        public static double GetBattleHpRatio()
        {
            double result = 0;
            int x, y;
            int xOffset_key = Coordinate.windowBoxLineOffset + 24;
            int yOffset_key = Coordinate.windowHOffset + 30;
            System.Drawing.Color HpExistColor;
            x = Coordinate.windowTop[0] + xOffset_key;
            y = Coordinate.windowTop[1] + yOffset_key;
            HpExistColor = System.Drawing.Color.FromArgb(254, 128, 128);
            using (Bitmap screenshot_Hp = BitmapFunction.CaptureScreen(x, y, 90, 1))
            {
                result = BitmapFunction.CalculateColorRatio(screenshot_Hp, HpExistColor);
            }
            return result;
        }
        public static bool checkCheatingCheck()
        {
            if (GameScript.isCheckCheatEnable == false) // 如果關閉此功能 , 直接回傳true
                return true;
            // 除檢查作弊畫面外 , 也可以是其他對話框
            int x_cheatCheck, y_cheatCheck;
            int xOffset_LU = Coordinate.windowBoxLineOffset + 270;
            int yOffset_LU = Coordinate.windowHOffset + 103;
            x_cheatCheck = Coordinate.windowTop[0] + xOffset_LU;
            y_cheatCheck = Coordinate.windowTop[1] + yOffset_LU;
            Bitmap cheatCheckBitmap;

            cheatCheckBitmap = Properties.Resources.Win7_CheatCheck__270_103_40x10;
            using (Bitmap screenshot_cheatCheckBitmap = BitmapFunction.CaptureScreen(x_cheatCheck, y_cheatCheck, 40, 10))
            {
                double Final_CCC = BitmapFunction.CompareImages(screenshot_cheatCheckBitmap, cheatCheckBitmap);
                Log.Append($"作弊檢測值為 '{Final_CCC}')\n");

                if (Final_CCC > 90)
                    return true;
                else
                    return false;
            }
        }
        public static bool GetBattleTimeEnd()
        {
            double result = 0;
            int x, y;
            int xOffset_key = Coordinate.windowBoxLineOffset + 23;
            int yOffset_key = Coordinate.windowHOffset + 7;
            System.Drawing.Color BTExistColor;
            x = Coordinate.windowTop[0] + xOffset_key;
            y = Coordinate.windowTop[1] + yOffset_key;
            BTExistColor = System.Drawing.Color.FromArgb(7, 70, 190);
            using (Bitmap screenshot_Hp = BitmapFunction.CaptureScreen(x, y, 160, 1))
            {
                result = BitmapFunction.CalculateColorRatio(screenshot_Hp, BTExistColor);
            }

            if (result == 0)
                return true;
            else
                return false;
        }
    }
}
