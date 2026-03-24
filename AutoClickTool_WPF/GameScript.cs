using System;
using System.Drawing;
using System.Threading;
using System.Windows;
using System.Windows.Input;

using AutoClickTool_WPF.Tool;

namespace AutoClickTool_WPF
{
    public class GameScript
    {
        // check cheat status
        public static bool isCheckCheatEnable = true;   // 檢查開關
        public static bool isCheckCheat = false;
        public static bool isMouseDefault = false;
        public static bool isDisplayRound = false;
        public static bool isPvpMode = false;
        // const value
        public const int enemyIndexMax = 9;
        public const int enemyIndexDefault = 0;
        public const int itemIndexMax = 15;
        public const int itemIndexDefault = 0;
        public const int roundNumberDefault = 0;        // 0 非戰鬥
                                                        // support value
        public static double mpRatioLess = 0.5;
        // using value
        public static int pollingEnemyIndex = enemyIndexDefault;
        public static int pollingWalkDirection = 0;
        public static int pollingItemIndex = itemIndexDefault;
        public static int pollingItemCheckTime = itemIndexDefault;
        public static int petSupTarget = 0;
        public static int buffTarget = 0;
        public static int usingItem = 0;
        public static int roundNumber = roundNumberDefault;
        public static int trainingSkillBattleType = 0;    // 0 = 隊員 , 1 = 召怪(暗使) , 2 = 自動走路
        public static int trainingSkillEndType = 0; // 0 = 技能 , 1 = 普通攻擊
        public static bool isSecondRoundSkill_Option = false;
        public static bool isSecondRoundSkill = false;
        public static int selectedItemIndex = 0;
        public static int selectedItemForm = 0;
        //Delay
        public static int autoWalkDelay = 200;
        public static int keyPressDelay = 50;
        public static int openItemDelay = 150;
        public static int keyPressDelayLong = 100;
        public static int summonDelay = 0;
        public static int spellDelay = 10;
        public static int spellRound = 1;

        public static int buyCBCount = 0;               // 當前購買數
        public static int buyCBCountTarget = 0;     // 指定購買數

        public static bool isClickRight = true;
        // Option Setting
        public static bool isItemRecheck = false;
        public static bool isItemEnd = false;
        public static bool isPetSupport = false;
        public static bool isPetSupportJust1 = false;
        public static bool isPetSupportToEnemy = false;
        public static bool isSummonerAttack = false;
        public static bool isLikeHuman = false;
        public static bool isAutoWalk = false;
        public static bool isAutoWalk_OptionRandom = false;
        public static bool isRoundLock = false;
        public static bool isTrainingSkillToEnemy = false;
        // Hotkey Setting
        public static Key playerActionKey = Key.F5;
        public static Key summonerActionKey = Key.F5;
        public static Key summonerAttackKey = Key.F6;
        public static Key petActionKey = Key.F5;
        public static Key trainingSkillKey = Key.F5;

        public static void mousePostionDefault()
        {
            int xOffset = Coordinate.windowBoxLineOffset + Coordinate.windowTop[0];
            int yOffset = Coordinate.windowHOffset + Coordinate.windowTop[1];
            // 移動到不會影響截圖判斷的位置
            if (isMouseDefault)
                MouseSimulator.MoveMouseTo(360 + xOffset, 580 + yOffset);
        }
        public static void setGameScriptFlagDefault()
        {
            // 作弊檢查狀態初始化
            isCheckCheat = false;
            // 腳本執行狀態初始化
            isPetSupport = false;
            isPetSupportJust1 = false;
            isPetSupportToEnemy = false;
            isSummonerAttack = false;
            isLikeHuman = false;
            isAutoWalk = false;
            isAutoWalk_OptionRandom = false;
            isSecondRoundSkill_Option = false;
            isRoundLock = false;
            isTrainingSkillToEnemy = false;
            // 腳本參數初始化
            pollingEnemyIndex = enemyIndexDefault;
            pollingEnemyIndex = enemyIndexDefault;
            pollingItemIndex = itemIndexDefault;
            pollingWalkDirection = 0;
        }
        public static void enemyPolling()
        {
            pollingEnemyIndex++;
            if (pollingEnemyIndex > enemyIndexMax)
                pollingEnemyIndex = enemyIndexDefault;
        }
        public static void itemPolling()
        {
            pollingItemIndex++;
            if (pollingItemIndex > itemIndexMax)
            {
                pollingItemIndex = itemIndexDefault;
                if (isItemRecheck)
                {
                    pollingItemCheckTime++;
                    GameFunction.pressRecheckBag();
                }
            }
        }
        public static void WalkDirectionPolling()
        {
            if (isAutoWalk_OptionRandom)
            {
                Random random = new Random();
                pollingWalkDirection = random.Next(0, 4);
            }
            else
            {
                pollingWalkDirection++;
                if (pollingWalkDirection > 3)
                    pollingWalkDirection = 0;
            }
        }
        public static void scriptOpenItemCB()
        {
            int xo, yo;
            if (GameFunction.CheckItemCB_Coor(pollingItemIndex, out xo, out yo))
            {
                pollingItemCheckTime = itemIndexDefault;
                // 開啟魔晶寶箱
                MouseSimulator.RightMousePress(xo, yo);
                Thread.Sleep(openItemDelay);
                // 結束後移開滑鼠避免誤判
                mousePostionDefault();
            }
            else
            {
                itemPolling();
                if (isItemRecheck)
                {
                    if (pollingItemCheckTime >= 3)
                    {
                        // 把自己關了
                        KeyboardSimulator.KeyPress(Key.F11);
                    }
                }
            }
        }
        public static void scriptOpenItemAP()
        {
            int xo, yo;
            int x_key, y_key;
            int xOffset_key = Coordinate.windowBoxLineOffset + 249;
            int yOffset_key = Coordinate.windowHOffset + 282;
            x_key = Coordinate.windowTop[0] + xOffset_key;
            y_key = Coordinate.windowTop[1] + yOffset_key;

            if (GameFunction.CheckItemAP_Coor(pollingItemIndex, out xo, out yo))
            {
                // 點擊使用
                MouseSimulator.RightMousePress(xo, yo);
                Thread.Sleep(openItemDelay);
                // 使用確認
                if (GameFunction.CheckItemAP_Yes())
                {
                    MouseSimulator.LeftMousePress(x_key, y_key);
                    Thread.Sleep(openItemDelay);
                }
                // 結束後移開滑鼠避免誤判
                mousePostionDefault();
                Thread.Sleep(openItemDelay);
            }
            else
            {
                itemPolling();
            }
        }
        public static void scriptNpcExchange_Alien()
        {
            // 401 323
            int xOffset = Coordinate.windowBoxLineOffset;
            int yOffset = Coordinate.windowHOffset;
            //CompareGameScreenshots
            Bitmap DialogsCheck = Properties.Resources.Win7_CheatCheck__270_103_40x10;
            Bitmap NpcAlien1 = Properties.Resources.win7_NpcExchangeAlien_1_270_280_;
            Bitmap NpcAlien2_option1 = Properties.Resources.win7_NpcExchangeAlien_2_option1_270_280_IntimacySugar;
            Bitmap NpcAlien2_option3 = Properties.Resources.win7_NpcExchangeAlien_2_option3_270_340_Silver;
            Bitmap NpcAlien3 = Properties.Resources.win7_NpcExchangeAlien_3_270_280_secondCheck;
            Bitmap NpcAlien4 = Properties.Resources.win7_NpcExchangeAlien_4_270_280_finalCheck;
            //// 檢查是否為對話框
            ///
            GameFunction.GameMoveMouseTo(0, 0);
            Thread.Sleep(keyPressDelay);
            if (BitmapFunction.CompareGameScreenshots(DialogsCheck, 270, 103, 100) == true)
            {
                // 檢查是否為我要兌換視窗
                if (BitmapFunction.CompareGameScreenshots(NpcAlien1, 270, 280, 100) == true)
                {
                    GameFunction.GameLeftMousePress(270, 280);
                    Thread.Sleep(500);
                    GameFunction.GameMoveMouseTo(0, 0);
                    Thread.Sleep(keyPressDelay);

                    if (BitmapFunction.CompareGameScreenshots(DialogsCheck, 270, 103, 100))
                    {
                        // 檢查是否為選擇獎勵 , 並檢查選項3
                        if (BitmapFunction.CompareGameScreenshots(NpcAlien2_option3, 270, 340, 100))
                        {
                            GameFunction.GameLeftMousePress(270, 340);
                            Thread.Sleep(500);
                            GameFunction.GameMoveMouseTo(0, 0);
                            Thread.Sleep(keyPressDelay);
                            if (BitmapFunction.CompareGameScreenshots(DialogsCheck, 270, 103, 100))
                            {
                                // 檢查是否進入到確認環節
                                if (BitmapFunction.CompareGameScreenshots(NpcAlien3, 270, 280, 100))
                                {
                                    GameFunction.GameLeftMousePress(270, 280);
                                    Thread.Sleep(500);
                                    GameFunction.GameMoveMouseTo(0, 0);
                                    Thread.Sleep(keyPressDelay);

                                    if (BitmapFunction.CompareGameScreenshots(DialogsCheck, 270, 103, 100))
                                    {
                                        // 檢查是否進入到結束後確認環節
                                        if (BitmapFunction.CompareGameScreenshots(NpcAlien4, 270, 280, 100))
                                        {
                                            GameFunction.GameLeftMousePress(270, 280);
                                            Thread.Sleep(500);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        public static void scriptNpcByeCB()
        {
            int xOffset = Coordinate.windowBoxLineOffset;
            int yOffset = Coordinate.windowHOffset;
            //CompareGameScreenshots
            Bitmap checkBuyBox = Properties.Resources.win7_x269_y172_BuyCBCheck;
            Bitmap checkBuyBoxCount = Properties.Resources.win7_x519_y189_CheckBuyCount;
            GameFunction.GameMoveMouseTo(0, 0);
            Thread.Sleep(keyPressDelay);
            if (BitmapFunction.CompareGameScreenshots(checkBuyBox, 269, 172, 100) == true)
            {
                // 檢查是否進入購買準備視窗
                GameFunction.GameLeftMousePress(498, 184);
                Thread.Sleep(100);
                GameFunction.GameRightMousePress(666, 251);
                if (BitmapFunction.CompareGameScreenshots(checkBuyBoxCount, 519, 189, 100) == true)
                {
                    // 檢查是否進入購買確認數量畫面
                    for (int i = 0; i < 5; i++)
                    {
                        KeyboardSimulator.KeyPress(Key.Back);
                        Thread.Sleep(100);
                    }
                    KeyboardSimulator.KeyPress(Key.D8);
                    Thread.Sleep(100);
                    KeyboardSimulator.KeyPress(Key.D0);
                    Thread.Sleep(100);
                    GameFunction.GameLeftMousePress(705, 301);
                    Thread.Sleep(200);
                    GameFunction.GameLeftMousePress(535, 431);

                    buyCBCount += 80;
                    if (buyCBCountTarget > 0 && buyCBCount >= buyCBCountTarget)
                    {
                        MessageBox.Show($"購買數已達到設定值");
                        KeyboardSimulator.KeyPress(Key.F11);
                        Thread.Sleep(500);
                    }
                }
                else
                {
                    MessageBox.Show($"抓取確認數量畫面失敗");
                    KeyboardSimulator.KeyPress(Key.F11);
                    Thread.Sleep(500);
                }
            }
        }
        public static void scriptAutoWalk()
        {
            if (isAutoWalk)
            {
                switch (pollingWalkDirection)
                {
                    case 0:
                        MouseSimulator.LeftMousePress(Coordinate.walkPoint[0, 0], Coordinate.walkPoint[0, 1]);
                        Thread.Sleep(autoWalkDelay);
                        break;
                    case 1:
                        MouseSimulator.LeftMousePress(Coordinate.walkPoint[1, 0], Coordinate.walkPoint[1, 1]);
                        Thread.Sleep(autoWalkDelay);
                        break;
                    case 2:
                        MouseSimulator.LeftMousePress(Coordinate.walkPoint[2, 0], Coordinate.walkPoint[2, 1]);
                        Thread.Sleep(autoWalkDelay);
                        break;
                    case 3:
                        MouseSimulator.LeftMousePress(Coordinate.walkPoint[3, 0], Coordinate.walkPoint[3, 1]);
                        Thread.Sleep(autoWalkDelay);
                        break;
                    default:
                        MouseSimulator.LeftMousePress(Coordinate.walkPoint[0, 0], Coordinate.walkPoint[0, 1]);
                        Thread.Sleep(autoWalkDelay);
                        break;
                }
            }
            WalkDirectionPolling();
        }

        //========================================
        public static void BattlaAction_SpellCastOnEnemy(Key keyCode)
        {
            int i = GameFunction.getEnemyCoor();
            if (i > 0)
            {
                GameFunction.castSpellOnTarget(Coordinate.Enemy[i - 1, 0], Coordinate.Enemy[i - 1, 1], keyCode, spellDelay);
                return;
            }
            GameFunction.castSpellOnTarget(Coordinate.Enemy[pollingEnemyIndex, 0], Coordinate.Enemy[pollingEnemyIndex, 1], keyCode, spellDelay);
            enemyPolling();
        }
        public static void BattlaAction_AttackOnEnemy()
        {
            int i = GameFunction.getEnemyCoor();
            if (i > 0)
            {
                GameFunction.AttackOnTarget(Coordinate.Enemy[i - 1, 0], Coordinate.Enemy[i - 1, 1]);
                return;
            }
            GameFunction.AttackOnTarget(Coordinate.Enemy[pollingEnemyIndex, 0], Coordinate.Enemy[pollingEnemyIndex, 1]);
            enemyPolling();
        }
        /*
         *  1. 計算時機在進入玩家回合的時候 , 計算後上Lock防止多算
         *  2. Lock在戰鬥計時器歸零時解鎖 , 檢查點在非人物、非幻獸視角
         *  3. 預計在練技輔助中使用 , 或以option開啟 , 回合數可以顯示在標題
         *  4. 預計可以加入場次計算
         */
        public static void BattleAction_PlayerRoundCheck()
        {
            if (isRoundLock == false)
            {
                isRoundLock = true;
                roundNumber++;
            }
        }
        public static void NormalAction_RoundCheck_True()
        {
            isRoundLock = false;
        }
        public static void NormalAction_RoundCheck_Reset()
        {
            isSecondRoundSkill = true;
            roundNumber = 0;
        }
        //==============BranchScript==============
        public static void BattleAction_Training()
        {
            double mp = GameFunction.GetBattleMpRatio();
            if (mp > mpRatioLess)
            {
                if (roundNumber > spellRound || roundNumber == 1)
                {
                    if (isTrainingSkillToEnemy == true)
                        BattlaAction_SpellCastOnEnemy(trainingSkillKey);
                    else
                        GameFunction.castSpellOnTarget(Coordinate.Friends[buffTarget, 0], Coordinate.Friends[buffTarget, 1], trainingSkillKey, spellDelay);

                    if (roundNumber > spellRound)
                        roundNumber = 1;
                }
                else if (isSecondRoundSkill_Option == true && isSecondRoundSkill == true && roundNumber == 2)
                {
                    isSecondRoundSkill = false;
                    GameFunction.castSpellOnTarget(Coordinate.Friends[2, 0], Coordinate.Friends[2, 1], Key.F5, spellDelay);
                }
                else
                {
                    GameFunction.pressDefendButton();
                }
            }
            else
            {
                switch (trainingSkillEndType)
                {
                    case 0:
                        BattlaAction_SpellCastOnEnemy(playerActionKey);
                        break;
                    case 1:
                    default:
                        BattlaAction_AttackOnEnemy();
                        break;
                }
            }
        }
        public static void BattleAction_SummonerAttack()
        {
            BattlaAction_SpellCastOnEnemy(summonerAttackKey);
        }
        public static void BattleAction_PlayerAttack()
        {
            BattlaAction_SpellCastOnEnemy(playerActionKey);
        }
        public static void BattleAction_Pet()
        {
            if (isPetSupport)
            {
                if (isPetSupportJust1 == true && roundNumber > 0)
                    GameFunction.pressDefendButton();
                else
                {
                    if (isPetSupportToEnemy)
                    {
                        BattlaAction_SpellCastOnEnemy(petActionKey);
                    }
                    else
                        GameFunction.castSpellOnTarget(Coordinate.Friends[petSupTarget, 0], Coordinate.Friends[petSupTarget, 1], petActionKey, spellDelay);
                }
            }
            else
            {
                GameFunction.pressDefendButton();
            }
        }
        public static void NormalAction_AutoWalk()
        {
            if (isAutoWalk == true)
            {
                if (GameFunction.NormalCheck() == true)
                {
                    if (GameFunction.checkCheatingCheck())
                    {
                        // 可能檢查到作弊檢測視窗 , 暫停所有動作
                        isCheckCheat = true;
                        return;
                    }
                    scriptAutoWalk();
                }
            }
            else
            {
                mousePostionDefault();
            }
        }
        public static void NormalAction_Summoner()
        {
            // 自動招怪
            if (GameFunction.NormalCheck() == true)
            {
                if (GameFunction.checkCheatingCheck())
                {
                    // 可能檢查到作弊檢測視窗 , 暫停所有動作
                    isCheckCheat = true;
                    return;
                }
                Thread.Sleep(summonDelay);  // 招怪前延遲
                KeyboardSimulator.KeyPress(summonerActionKey);
                Thread.Sleep(keyPressDelayLong); // 多一點延遲
            }
        }
        public static void NormalAction_Training()
        {
            switch (trainingSkillBattleType)
            {
                case 0:
                    break;
                case 1:
                    NormalAction_Summoner();
                    break;
                case 2:
                    NormalAction_AutoWalk();
                    break;
                default:
                    break;
            }
        }
        //==============MainScript==============
        public static void AutoBattle()
        {
            if (GameFunction.BattleCheck_Player() == true)
            {
                BattleAction_PlayerRoundCheck();

                BattleAction_PlayerAttack();
            }
            else if (GameFunction.BattleCheck_Pet() == true)
            {
                BattleAction_Pet();
            }
            else
            {
                if (GameFunction.NormalCheck() == true)
                {
                    NormalAction_RoundCheck_Reset();
                }
                else
                {
                    NormalAction_RoundCheck_True();
                }
                NormalAction_AutoWalk();

                if (isPvpMode == true)
                {
                    KeyboardSimulator.KeyPress(Key.F4);
                    Thread.Sleep(200);
                    KeyboardSimulator.KeyPress(Key.F4);
                    Thread.Sleep(200);
                }
            }
        }
        public static void AutoDefend()
        {
            if (GameFunction.BattleCheck_Player() == true)
            {
                BattleAction_PlayerRoundCheck();

                GameFunction.pressDefendButton();
            }
            else if (GameFunction.BattleCheck_Pet() == true)
            {
                BattleAction_Pet();
            }
            else
            {
                if (GameFunction.NormalCheck() == true)
                {
                    NormalAction_RoundCheck_Reset();
                }
                else
                {
                    NormalAction_RoundCheck_True();
                }

                NormalAction_AutoWalk();
            }
        }
        public static void AutoEnterBattle()
        {
            if (GameFunction.BattleCheck_Player() == true)
            {
                BattleAction_PlayerRoundCheck();

                if (isSummonerAttack)
                {
                    BattleAction_SummonerAttack();
                }
                else
                    GameFunction.pressDefendButton();
            }
            else if (GameFunction.BattleCheck_Pet() == true)
            {
                BattleAction_Pet();
            }
            else
            {
                if (GameFunction.NormalCheck() == true)
                {
                    NormalAction_RoundCheck_Reset();
                }
                else
                {
                    NormalAction_RoundCheck_True();
                }

                NormalAction_Summoner();
            }
        }
        public static void AutoBuff()
        {
            if (GameFunction.BattleCheck_Player() == true)
            {
                BattleAction_PlayerRoundCheck();

                GameFunction.castSpellOnTarget(Coordinate.Friends[buffTarget, 0], Coordinate.Friends[buffTarget, 1], playerActionKey, spellDelay);
            }
            else if (GameFunction.BattleCheck_Pet() == true)
            {
                BattleAction_Pet();
            }
            else
            {
                if (GameFunction.NormalCheck() == true)
                {
                    NormalAction_RoundCheck_Reset();
                }
                else
                {
                    NormalAction_RoundCheck_True();
                }

                NormalAction_AutoWalk();
            }
        }
        public static void AutoUsingItem()
        {
            if (usingItem == 1 || usingItem == 2)
            {
                if (GameFunction.NormalCheck() == true)
                {
                    if (GameFunction.ItemTimeCheck())
                    {
                        switch (usingItem)
                        {
                            case 1:
                                // 自動開啟魔晶箱
                                scriptOpenItemCB();
                                break;
                            case 2:
                                // 自動開啟調整藥丸
                                scriptOpenItemAP();
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        // 打開物品欄
                        GameFunction.openToralItemScene();
                        Thread.Sleep(keyPressDelay);
                    }
                }
            }
            else if (usingItem == 3)
            {
                int xOffset = Coordinate.windowBoxLineOffset;
                int yOffset = Coordinate.windowHOffset;
                if (GameFunction.NormalCheck())
                {
                    MouseSimulator.RightMousePress(401 + xOffset, 315 + yOffset);
                    Thread.Sleep(1000);
                    scriptNpcExchange_Alien();
                }
                else
                {
                }
            }
            else if (usingItem == 4)
            {
                // 自動購買魔晶箱
                int xOffset = Coordinate.windowBoxLineOffset;
                int yOffset = Coordinate.windowHOffset;
                if (GameFunction.NormalCheck())
                {
                    KeyboardSimulator.KeyPress(Key.F4);
                    Thread.Sleep(200);
                    KeyboardSimulator.KeyPress(Key.F4);
                    Thread.Sleep(200);
                    MouseSimulator.RightMousePress(401 + xOffset, 315 + yOffset);
                    Thread.Sleep(500);
                    scriptNpcByeCB();
                }
                else
                {
                }
            }
            else if (usingItem == 5)
            {
                // 自動拉霸機.自動購買月餅
                int xOffset = Coordinate.windowBoxLineOffset;
                int yOffset = Coordinate.windowHOffset;

                if (isClickRight)
                {
                    isClickRight = false;
                    MouseSimulator.RightMousePress(404 + xOffset, 289 + yOffset);
                }
                else
                {
                    isClickRight = true;
                    MouseSimulator.LeftMousePress(404 + xOffset, 289 + yOffset);
                }
            }
            else if (usingItem == 6)
            {
                // 交易自動選取物品
                int xOffset = Coordinate.windowBoxLineOffset;
                int yOffset = Coordinate.windowHOffset;
                //檢查是否在交易畫面
                Bitmap TradeCheck = Properties.Resources.TradeCheck__x368_y128_;
                if (BitmapFunction.CompareGameScreenshots(TradeCheck, 368, 128, 100) == true)
                {
                    int page;
                    int pageMax = 5;
                    int itemIndex;
                    int itemTradeMax = 20;
                    int ErrorTime = 0;
                    //逐一檢查物品內容
                    for (page = 0; page < pageMax; page++)
                    {
                        for (itemIndex = 0; itemIndex < itemTradeMax; itemIndex++)
                        {
                            // 取得第 itemIndex 個物品座標
                            int xItem, yItem;
                            int xSelect, ySelect;
                            GameFunction.GetTradeItemCoor(itemIndex, out xItem, out yItem);
                            GameFunction.GetTradeItemSelectCoor(itemIndex, out xSelect, out ySelect);

                            Bitmap itemCheck = Properties.Resources.Item_x441_y199_DreamOre;
                            switch (selectedItemIndex)
                            {
                                case 1:
                                    itemCheck = Properties.Resources.Item_x441_y199_DreamOre;
                                    break;
                                case 2:
                                    itemCheck = Properties.Resources.Item_x441_y199_FillolOre;
                                    break;
                                case 3:
                                    itemCheck = Properties.Resources.Item_x441_y199_Emery;
                                    break;
                                default:
                                    break;
                            }
                            // 檢查是否為物品
                            Log.Append($"檢查是否為物品 , PAGE=[{page}] , ITEM = [{itemIndex}], X=[{xItem}],Y=[{yItem}]");
                            if (selectedItemIndex == 0 || BitmapFunction.CompareGameScreenshots(itemCheck, xItem, yItem, 100) == true)
                            {
                                // 是否為未選取
                                Bitmap notSelectCheck = Properties.Resources.TradeNotSelect__x425_y184_;
                                if (BitmapFunction.CompareGameScreenshots(notSelectCheck, xSelect, ySelect, 100) == true)
                                {
                                    MouseSimulator.RightMousePress(xItem + xOffset, yItem + yOffset);
                                    Thread.Sleep(openItemDelay);
                                    Log.Append($"物品選取按下");
                                    Bitmap selectCheck = Properties.Resources.TradeIsSelect__x425_y184_;
                                    // 檢查是否為選中
                                    Log.Append($"檢查是否選取 , PAGE=[{page}] , ITEM = [{itemIndex}], X=[{xSelect}],Y=[{ySelect}]");
                                    if (BitmapFunction.CompareGameScreenshots(selectCheck, xSelect, ySelect, 100) == true)
                                    {
                                        Log.Append($"物品選已選取");
                                    }
                                    else
                                    {
                                        Log.Append($"物品未選取");
                                    }
                                }
                            }
                            else
                            {
                            }
                        }
                        // 按下一頁
                        MouseSimulator.LeftMousePress(498 + xOffset, 442 + yOffset);
                    }
                    KeyboardSimulator.KeyPress(Key.F11);
                }
                else
                {
                    KeyboardSimulator.KeyPress(Key.F11);
                }
            }
            else if (usingItem == 7)
            {
                // 交易自動選取物品
                int xOffset = Coordinate.windowBoxLineOffset;
                int yOffset = Coordinate.windowHOffset;
                // 取得第 itemIndex 個物品座標
                int xItem, yItem;
                int xSelect, ySelect;
                int itemIndex = 0;
                //檢查是否在交易畫面
                Bitmap InventoryBankCheck = Properties.Resources.InventoryAndBank_x196_y303_;
                if (BitmapFunction.CompareGameScreenshots(InventoryBankCheck, 196, 303, 100) == true)
                {
                    int itemTradeMax;
                    // 是否為未選取
                    Bitmap notSelectCheck;
                    Bitmap selectCheck;
                    if (selectedItemForm == 0)
                    {
                        // 物品欄 起點528 337 5x5  X57Y51
                        itemTradeMax = 16;
                        selectCheck = Properties.Resources.x555_y331_InventorySelected;
                    }
                    else
                    {
                        // 銀行 起點221 96 5x5  X45Y44
                        itemTradeMax = 40;
                        selectCheck = Properties.Resources.x248_y87_BankSelected;
                    }

                    Bitmap itemCheck = Properties.Resources.Item_x441_y199_DreamOre;
                    switch (selectedItemIndex)
                    {
                        case 1:
                            itemCheck = Properties.Resources.Item_x441_y199_Emery;
                            break;
                        default:
                            break;
                    }


                    for (itemIndex = 0; itemIndex < itemTradeMax; itemIndex++)
                    {
                        if (selectedItemForm == 0)
                        {
                            GameFunction.GetInventoryItemCoor(itemIndex, out xItem, out yItem);
                            GameFunction.GetInventoryItemSelectedCoor(itemIndex, out xSelect, out ySelect);
                        }
                        else
                        {
                            GameFunction.GetBankItemCoor(itemIndex, out xItem, out yItem);
                            GameFunction.GetBankItemSelectedCoor(itemIndex, out xSelect, out ySelect);
                        }

                        if (selectedItemIndex == 0 || BitmapFunction.CompareGameScreenshots(itemCheck, xItem, yItem, 100) == true)
                        {
                            MouseSimulator.RightMousePress(xItem + xOffset, yItem + yOffset);
                            Thread.Sleep(openItemDelay);
                            if (BitmapFunction.CompareGameScreenshots(selectCheck, xSelect, ySelect, 100) == true)
                            {
                                Log.Append($"物品選已選取");
                            }
                            else
                            {
                                // 物品欄會因為雜點導致判斷失敗故直接不判斷, 先點一次若非選取則補選取
                                MouseSimulator.RightMousePress(xItem + xOffset, yItem + yOffset);
                                Thread.Sleep(openItemDelay);
                            }
                        }
                    }
                    MessageBox.Show($"完成");
                    KeyboardSimulator.KeyPress(Key.F11);
                }
                else
                {
                    MessageBox.Show($"銀行未進入");
                    KeyboardSimulator.KeyPress(Key.F11);
                }
            }
        }
        public static void AutoSkillTraining()
        {
            if (GameFunction.BattleCheck_Player() == true)
            {
                BattleAction_PlayerRoundCheck();

                BattleAction_Training();
            }
            else if (GameFunction.BattleCheck_Pet() == true)
            {
                BattleAction_Pet();
            }
            else
            {
                if (GameFunction.NormalCheck() == true)
                {
                    NormalAction_RoundCheck_Reset();
                }
                else
                {
                    NormalAction_RoundCheck_True();
                }

                NormalAction_Training();
            }
        }
    }
}
