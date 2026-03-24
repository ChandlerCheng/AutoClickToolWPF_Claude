using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoClickTool_WPF.Tool
{
    public class Coordinate
    {
        // 視窗資訊
        public static int[] windowTop = new int[2];
        public static int[] windowBottom = new int[2];
        public static int windowHeigh = 0;
        public static int windowWidth = 0;
        public static int windowHOffset = 30;
        public static int windowBoxLineOffset = 8;
        public static bool IsGetWindows = false;
        /*
                敵方座標(目視)
                67890
                12345
                
                陣列值
                56789
                01234
         */
        public static int[,] Enemy = new int[10, 2];
        public static int[,] checkEnemy = new int[10, 2];
        public static int[,] checkEnemy_2 = new int[10, 2];
        public static int[,] checkEnemy_3 = new int[10, 2];
        public static int[,] walkPoint = new int[4, 2]; // 上下左右
        /*
               我方座標(目視)
               12345
               67890

               陣列值
               01234
               56789
        */
        public static int[,] Friends = new int[10, 2];
        public static void CalculateWalkPoint()
        {
            int xOffset = windowBoxLineOffset + windowTop[0];
            int yOffset = windowHOffset + +windowTop[1];
            /*
                上 411 ,48
                下 411,479
                左 102,257
                右 652,261
             */

            // 上
            walkPoint[0, 0] = 411 + xOffset;
            walkPoint[0, 1] = 48 + yOffset;

            //下
            walkPoint[1, 0] = 411 + xOffset;
            walkPoint[1, 1] = 479 + yOffset;

            //左
            walkPoint[2, 0] = 102 + xOffset;
            walkPoint[2, 1] = 257 + yOffset;

            // 右
            walkPoint[3, 0] = 652 + xOffset;
            walkPoint[3, 1] = 257 + yOffset;
        }
        public static void CalculateAllEnemy(int x, int y)
        {
            // 計算第三號敵人的座標 (陣列索引為2)
            Enemy[2, 0] = x;
            Enemy[2, 1] = y;

            // 計算其他敵人的座標
            CalculateTargetCoordinate(Enemy, 2, 1, -68, 56);
            CalculateTargetCoordinate(Enemy, 1, 0, -68, 56);
            CalculateTargetCoordinate(Enemy, 2, 3, 68, -56);
            CalculateTargetCoordinate(Enemy, 3, 4, 68, -56);

            CalculateTargetCoordinate(Enemy, 0, 5, -73, -61);
            CalculateTargetCoordinate(Enemy, 1, 6, -73, -61);
            CalculateTargetCoordinate(Enemy, 2, 7, -73, -61);
            CalculateTargetCoordinate(Enemy, 3, 8, -73, -61);
            CalculateTargetCoordinate(Enemy, 4, 9, -73, -61);
        }
        public static void CalculateAllFriends(int x, int y)
        {
            // 計算第三號敵人的座標 (陣列索引為2)
            Friends[2, 0] = x;
            Friends[2, 1] = y;
            // 計算其他的座標
            CalculateTargetCoordinate(Friends, 2, 1, -68, 56);
            CalculateTargetCoordinate(Friends, 1, 0, -68, 56);
            CalculateTargetCoordinate(Friends, 2, 3, 68, -56);
            CalculateTargetCoordinate(Friends, 3, 4, 68, -56);

            CalculateTargetCoordinate(Friends, 0, 5, 73, 61);
            CalculateTargetCoordinate(Friends, 1, 6, 73, 61);
            CalculateTargetCoordinate(Friends, 2, 7, 73, 61);
            CalculateTargetCoordinate(Friends, 3, 8, 73, 61);
            CalculateTargetCoordinate(Friends, 4, 9, 73, 61);
        }
        public static void CalculateEnemyCheckXY()
        {
            /*
             * 抓取是否有怪物的白色圓環 , 參考最大體型生物為威奇迷宮
             *  20240424 - > 實際上不可行 , 上半截怪物體型大會被遮住(且怪物有循環動作)
             *  下半截則在1號位會有文字遮住怪物圓環的狀況
             */
            checkEnemy[0, 0] = 100;
            checkEnemy[0, 1] = 399;

            checkEnemy_2[0, 0] = 128;
            checkEnemy_2[0, 1] = 423;
            // 第一檢查點 , 土萊姆會失效
            CalculateTargetCoordinate(checkEnemy, 0, 1, 72, -54);
            CalculateTargetCoordinate(checkEnemy, 1, 2, 72, -54);
            CalculateTargetCoordinate(checkEnemy, 2, 3, 72, -54);
            CalculateTargetCoordinate(checkEnemy, 3, 4, 72, -54);

            CalculateTargetCoordinate(checkEnemy, 0, 5, -64, -60);
            CalculateTargetCoordinate(checkEnemy, 1, 6, -64, -60);
            CalculateTargetCoordinate(checkEnemy, 2, 7, -64, -60);
            CalculateTargetCoordinate(checkEnemy, 3, 8, -64, -60);
            CalculateTargetCoordinate(checkEnemy, 4, 9, -64, -60);

            //第二檢查點 , 萊姆系在第一位會失效(被文字遮住) , 但在第二位沒問題
            CalculateTargetCoordinate(checkEnemy_2, 0, 1, 72, -54);
            CalculateTargetCoordinate(checkEnemy_2, 1, 2, 72, -54);
            CalculateTargetCoordinate(checkEnemy_2, 2, 3, 72, -54);
            CalculateTargetCoordinate(checkEnemy_2, 3, 4, 72, -54);

            CalculateTargetCoordinate(checkEnemy_2, 0, 5, -64, -60);
            CalculateTargetCoordinate(checkEnemy_2, 1, 6, -64, -60);
            CalculateTargetCoordinate(checkEnemy_2, 2, 7, -64, -60);
            CalculateTargetCoordinate(checkEnemy_2, 3, 8, -64, -60);
            CalculateTargetCoordinate(checkEnemy_2, 4, 9, -64, -60);
        }
        // 計算目標的座標
        private static void CalculateTargetCoordinate(int[,] enemyArray, int fromIndex, int toIndex, int xOffset, int yOffset)
        {
            // 获取源坐标的值
            int fromX = enemyArray[fromIndex, 0];
            int fromY = enemyArray[fromIndex, 1];

            // 计算目标坐标并赋值给目标索引
            enemyArray[toIndex, 0] = fromX + xOffset;
            enemyArray[toIndex, 1] = fromY + yOffset;
        }
    }
}
