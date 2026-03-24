using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Interop;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using AutoClickTool_WPF.Tool;

namespace AutoClickTool_WPF
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        private static Task actionTask;
        private static bool isRelease = false;      // 顯示發行版或測試版
        private static bool isOnlySupport = false;      // 顯示發行版或測試版
        private static bool isEnable = false;       // 程式標題狀態更改
        private static bool isWindowLoaded = false;     // 用於檢查畫面上物件是否已經存在
        private static int tabFunctionSelected = 0;     // 腳本號碼
        private static bool isChangeTitle = false;      // 作弊檢查的標題更改
        private static CancellationTokenSource cancellationTokenSource;
        #region 取得系統資訊
        // 定義結構來存儲版本資訊
        [StructLayout(LayoutKind.Sequential)]
        public struct OSVERSIONINFOEX
        {
            public uint dwOSVersionInfoSize;
            public uint dwMajorVersion;
            public uint dwMinorVersion;
            public uint dwBuildNumber;
            public uint dwPlatformId;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
            public string szCSDVersion;
        }

        // 使用 P/Invoke 調用 RtlGetVersion
        [DllImport("ntdll.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int RtlGetVersion(ref OSVERSIONINFOEX versionInfo);
        #endregion
        #region 程式初始化
        public MainWindow()
        {
            InitializeComponent();
        }
        protected override void OnSourceInitialized(EventArgs e)
        {
            var comboItems = new List<string> { "F5", "F6", "F7", "F8", "F9", "F10", "F12" };
            var comboItems_Debug = new List<string> { "F5", "F6", "F7", "F8", "F10", "F11", "F12" };
            var comboItemsRatio = new List<string> { "10%", "20%", "30%", "40%", "50%", "60%", "70%", "80%", "90%" };

            base.OnSourceInitialized(e);
            #region WINDOWS檢查
            OSVERSIONINFOEX osVersion = new OSVERSIONINFOEX();
            RtlGetVersion(ref osVersion);
            string version = $"{osVersion.dwMajorVersion}.{osVersion.dwMinorVersion}.{osVersion.dwBuildNumber}";
            if (osVersion.dwMajorVersion == 10)
                SystemSetting.isWin10 = true;
            #endregion
            #region 畫面DEBUG檢查
#if DEBUG
            SystemSetting.isDebug = true;
            this.tabIndexPage.Visibility = Visibility.Collapsed;
            tabControlUsingMethod.SelectedIndex = 1;
#else
            this.tabTestFunction.Visibility = Visibility.Collapsed;
            tabControlUsingMethod.SelectedIndex = 0;
#endif
            #endregion
            #region 自定義功能顯示控制

            // =============  純輔助功能 =================
            // 隱藏多餘選項
            if (isOnlySupport == true)
            {
                tabAutoBattle.Visibility = Visibility.Collapsed;
                tabAutoBuff.Visibility = Visibility.Collapsed;
                tabAutoDefend.Visibility = Visibility.Collapsed;
                tabAutoEnterBattle.Visibility = Visibility.Collapsed;
                tabAutoBattle.Visibility = Visibility.Collapsed;
                tabSkillTraining.Visibility = Visibility.Collapsed;
                tabSetting.Visibility = Visibility.Collapsed;
                labelCheatCheckTitle.Visibility = Visibility.Collapsed;
                labelCheatCheckStatus.Visibility = Visibility.Collapsed;
                checkMouseDefault.Visibility = Visibility.Collapsed;
            }
            // 輔助功能未完成項
            radioAutoNpcExchangeAlien.Visibility = Visibility.Collapsed;
            // ============= ==========================

            #endregion
            // 此為初始化怪物檢查點 , 務必在視窗打開時優先動作
            Coordinate.CalculateEnemyCheckXY();
            var helper = new WindowInteropHelper(this);
            IntPtr hwnd = helper.Handle;
#if !DEBUG
            SystemSetting.RegisterHotKey(hwnd, SystemSetting.HOTKEY_SCRIPT_EN, 0, (uint)KeyInterop.VirtualKeyFromKey(Key.F11));
#else
            // DEBUG模式下 , 使用F11做熱鍵會干擾編譯時逐步執行
            SystemSetting.RegisterHotKey(hwnd, SystemSetting.HOTKEY_SCRIPT_EN, 0, (uint)KeyInterop.VirtualKeyFromKey(Key.F9));
            DebugFunction.IsDebugMsg = true;
#endif
            HwndSource source = HwndSource.FromHwnd(hwnd);
            source.AddHook(WndProc);
#if DEBUG
            labelHotkeyDisplay.Content = "F9";
            labelHotkeyDisplay.Foreground = new SolidColorBrush(System.Windows.Media.Colors.Red);
            labelBuildType.Content = "Debug";
            labelBuildType.Foreground = new SolidColorBrush(System.Windows.Media.Colors.Red);

            foreach (var item in comboItems_Debug)
            {
                tab2ComboPlayerActionKey.Items.Add(item);
                tab2ComboPetActionKey.Items.Add(item);
                tab3ComboPetActionKey.Items.Add(item);
                tab4ComboSummonerActionKey.Items.Add(item);
                tab4ComboSummonerAttackKey.Items.Add(item);
                tab4ComboPetActionKey.Items.Add(item);
                tab5ComboAutoBuffKey.Items.Add(item);
                tab5ComboPetActionKey.Items.Add(item);
                tab7ComboSummonerActionKey.Items.Add(item);
                tab7ComboSkillTrainingKey.Items.Add(item);
                tab7ComboFinalAttackKey.Items.Add(item);
            }
#else
            labelHotkeyDisplay.Content = "F11";
            labelHotkeyDisplay.Foreground = new SolidColorBrush(System.Windows.Media.Colors.Green);

            // 發行版或測試版
            if (isRelease)
            {
                labelBuildType.Content = "Release";
                labelBuildType.Foreground = new SolidColorBrush(System.Windows.Media.Colors.Green);
            }
            else
            {
                labelBuildType.Content = "Beta";
                labelBuildType.Foreground = new SolidColorBrush(System.Windows.Media.Colors.Red);
            }

            foreach (var item in comboItems)
            {
                tab2ComboPlayerActionKey.Items.Add(item);
                tab2ComboPetActionKey.Items.Add(item);
                tab3ComboPetActionKey.Items.Add(item);
                tab4ComboSummonerActionKey.Items.Add(item);
                tab4ComboSummonerAttackKey.Items.Add(item);
                tab4ComboPetActionKey.Items.Add(item);
                tab5ComboAutoBuffKey.Items.Add(item);
                tab5ComboPetActionKey.Items.Add(item);
                tab7ComboSummonerActionKey.Items.Add(item);
                tab7ComboSkillTrainingKey.Items.Add(item);
                tab7ComboFinalAttackKey.Items.Add(item);
                tab7ComboPetActionKey.Items.Add(item);
            }
#endif
            foreach (var item in comboItemsRatio)
                tab7ComboMPLessThan.Items.Add(item);

            if (GameScript.isCheckCheatEnable == true)
            {
                labelCheatCheckStatus.Content = Application.Current.Resources["commanEnable"].ToString();
                labelCheatCheckStatus.Foreground = new SolidColorBrush(System.Windows.Media.Colors.Green);
            }
            else
            {
                labelCheatCheckStatus.Content = Application.Current.Resources["commanDisable"].ToString();
                labelCheatCheckStatus.Foreground = new SolidColorBrush(System.Windows.Media.Colors.Red);
            }
            // 數值初始化並更新到設定中
            GameScript.openItemDelay = Properties.Settings.Default.itemDelay;
            textSupportDelay.Text = GameScript.openItemDelay.ToString();

            GameScript.keyPressDelay = Properties.Settings.Default.battleDelay;
            textKeyPressDelay.Text = GameScript.keyPressDelay.ToString();

            GameScript.keyPressDelayLong = Properties.Settings.Default.battleDelayLong;
            textMousePressDelay.Text = GameScript.keyPressDelayLong.ToString();

            GameScript.summonDelay = Properties.Settings.Default.summonDelay;
            textSummonDelay.Text = GameScript.summonDelay.ToString();

            // Log紀錄初始化
            Log.Init();
            isWindowLoaded = true;
        }
        #endregion
        #region 畫面更新
        private void ViewUpdate()
        {
            if (isWindowLoaded == false)
                return;

            Dispatcher.Invoke(() =>
            {
                if (GameScript.isCheckCheatEnable)
                {
                    labelCheatCheckStatus.Content = Application.Current.Resources["commanEnable"].ToString();
                    labelCheatCheckStatus.Foreground = new SolidColorBrush(System.Windows.Media.Colors.Green);
                }
                else
                {
                    labelCheatCheckStatus.Content = Application.Current.Resources["commanDisable"].ToString();
                    labelCheatCheckStatus.Foreground = new SolidColorBrush(System.Windows.Media.Colors.Red);
                }
            });
        }
        #endregion
        #region 遊戲中動作
        private static void BattleLoop()
        {
            // 偵測到作弊後 , 持續檢查畫面是否消除 , 是則自動重啟腳本
            if (GameScript.isCheckCheat == true)
            {
                // 由於超過三次帳號會被鎖定 , 故每次抓取圖片前都重新定位一次視窗位置
                SystemSetting.GetGameWindow();
                if (GameFunction.checkCheatingCheck() == false)
                {
                    GameScript.isCheckCheat = false;
                    isChangeTitle = false;
                    isEnable = true;
                    return;
                }
                return;
            }

            switch (tabFunctionSelected)
            {
                case 2:
                    GameScript.AutoBattle();
                    break;
                case 3:
                    GameScript.AutoDefend();
                    break;
                case 4:
                    GameScript.AutoEnterBattle();
                    break;
                case 5:
                    GameScript.AutoBuff();
                    break;
                case 6:
                    GameScript.AutoUsingItem();
                    break;
                case 7:
                    GameScript.AutoSkillTraining();
                    break;
                default:
                    break;
            }

        }
        #endregion
        #region 執行緒動作
        private static async Task ActionLoop(CancellationToken token, MainWindow window)
        {
            while (!token.IsCancellationRequested)
            {
                if (isEnable)
                {
                    // 更新視窗的 Title
                    window.Dispatcher.Invoke(() =>
                    {
                        // 單純想讓Title可以隨語系改變好看用
                        window.Title = Application.Current.Resources["windowTitleFunction"].ToString() + " '" + DebugFunction.feedBackFunctionTitle(tabFunctionSelected) +
                        "'" + " " + Application.Current.Resources["windowTitleIs"].ToString() + " " +
                        Application.Current.Resources["windowTitleRuning"].ToString();
                    });
                    isEnable = false;
                }
                // 更新並顯示場次於標題 , 使用於練技模式
                if (tabFunctionSelected == 7 || GameScript.isDisplayRound == true)
                {
                    window.Dispatcher.Invoke(() =>
                    {
                        window.Title = "" + GameScript.roundNumber + "回合";
                    });
                }
                else if (tabFunctionSelected == 6 && GameScript.usingItem == 4)
                {
                    window.Dispatcher.Invoke(() =>
                    {
                        window.Title = "" + GameScript.buyCBCount + "/" + GameScript.buyCBCountTarget;
                    });
                }
                // 檢測到外掛檢測畫面 , 變更標題
                if (GameScript.isCheckCheat == true && isChangeTitle == false)
                {
                    window.Dispatcher.Invoke(() =>
                    {
                        window.Title = Application.Current.Resources["windowTitleCheckCheat"].ToString();
                    });
                    isChangeTitle = true;
                }

                await Task.Delay(100);  // 模擬一些延遲，避免CPU過度佔用
                BattleLoop();
            }
            // 更新視窗的 Title
            window.Dispatcher.Invoke(() =>
                {
                    window.Title = Application.Current.Resources["windowTitleSuspending"].ToString();
                });
            isEnable = false;
        }
        public static async void HotKeyAction_Script_EnableSwitch(MainWindow window)
        {

            SystemSetting.GetGameWindow();
#if !DEBUG
            if (Coordinate.IsGetWindows != true)
            {
                MessageBox.Show(Application.Current.Resources["msgDebugGameTitleIsError"].ToString());

                return;
            }
#endif
            // 檢查是否已經有一個執行中的任務
            if (actionTask != null && !actionTask.IsCompleted)
            {
                cancellationTokenSource.Cancel();  // 發送取消請求
                try
                {
                    await actionTask;  // 等待任務結束
                }
                catch (OperationCanceledException)
                {
                    // 任務被取消的情況
                }
                finally
                {
                    cancellationTokenSource.Dispose();  // 清理
                }
                return;
            }

            GameScript.roundNumber = 0;

            // 檢查是否需要重新啟動
            if (actionTask == null || actionTask.IsCompleted)
            {
                cancellationTokenSource = new CancellationTokenSource();
                actionTask = Task.Run(() => ActionLoop(cancellationTokenSource.Token, window));
                isEnable = true;
            }
        }
        #endregion
        #region 資訊設定
        public static void setPlayerActionKey(int selectedIndex)
        {
            switch (selectedIndex)
            {
                case 0:
                    GameScript.playerActionKey = Key.F5;
                    break;
                case 1:
                    GameScript.playerActionKey = Key.F6;
                    break;
                case 2:
                    GameScript.playerActionKey = Key.F7;
                    break;
                case 3:
                    GameScript.playerActionKey = Key.F8;
                    break;
#if DEBUG
                case 4:
                    GameScript.playerActionKey = Key.F10;
                    break;
                case 5:
                    GameScript.playerActionKey = Key.F11;
                    break;
#else
                case 4:
                    GameScript.playerActionKey = Key.F9;
                    break;
                case 5:
                    GameScript.playerActionKey = Key.F10;
                    break;
#endif
                case 6:
                    GameScript.playerActionKey = Key.F12;
                    break;
                default:
                    GameScript.playerActionKey = Key.F5;
                    break;
            }
        }
        public static void setSummonerActionKey(int selectedIndex)
        {
            switch (selectedIndex)
            {
                case 0:
                    GameScript.summonerActionKey = Key.F5;
                    break;
                case 1:
                    GameScript.summonerActionKey = Key.F6;
                    break;
                case 2:
                    GameScript.summonerActionKey = Key.F7;
                    break;
                case 3:
                    GameScript.summonerActionKey = Key.F8;
                    break;
#if DEBUG
                case 4:
                    GameScript.summonerActionKey = Key.F10;
                    break;
                case 5:
                    GameScript.summonerActionKey = Key.F11;
                    break;
#else
                case 4:
                    GameScript.summonerActionKey = Key.F9;
                    break;
                case 5:
                    GameScript.summonerActionKey = Key.F10;
                    break;
#endif
                case 6:
                    GameScript.summonerActionKey = Key.F12;
                    break;
                default:
                    GameScript.summonerActionKey = Key.F5;
                    break;
            }
        }
        public static void setSummonerAttackKey(int selectedIndex)
        {
            switch (selectedIndex)
            {
                case 0:
                    GameScript.summonerAttackKey = Key.F5;
                    break;
                case 1:
                    GameScript.summonerAttackKey = Key.F6;
                    break;
                case 2:
                    GameScript.summonerAttackKey = Key.F7;
                    break;
                case 3:
                    GameScript.summonerAttackKey = Key.F8;
                    break;
#if DEBUG
                case 4:
                    GameScript.summonerAttackKey = Key.F10;
                    break;
                case 5:
                    GameScript.summonerAttackKey = Key.F11;
                    break;
#else
                case 4:
                    GameScript.summonerAttackKey = Key.F9;
                    break;
                case 5:
                    GameScript.summonerAttackKey = Key.F10;
                    break;
#endif
                case 6:
                    GameScript.summonerAttackKey = Key.F12;
                    break;
                default:
                    GameScript.summonerAttackKey = Key.F5;
                    break;
            }
        }
        public static void setPetActionKey(int selectedIndex)
        {
            switch (selectedIndex)
            {
                case 0:
                    GameScript.petActionKey = Key.F5;
                    break;
                case 1:
                    GameScript.petActionKey = Key.F6;
                    break;
                case 2:
                    GameScript.petActionKey = Key.F7;
                    break;
                case 3:
                    GameScript.petActionKey = Key.F8;
                    break;
#if DEBUG
                case 4:
                    GameScript.petActionKey = Key.F10;
                    break;
                case 5:
                    GameScript.petActionKey = Key.F11;
                    break;
#else
                case 4:
                    GameScript.petActionKey = Key.F9;
                    break;
                case 5:
                    GameScript.petActionKey = Key.F10;
                    break;
#endif
                case 6:
                    GameScript.petActionKey = Key.F12;
                    break;
                default:
                    break;
            }
        }
        public static void setTrainingSkillKey(int selectedIndex, int spellRound)
        {
            GameScript.spellRound = spellRound;
            switch (selectedIndex)
            {
                case 0:
                    GameScript.trainingSkillKey = Key.F5;
                    break;
                case 1:
                    GameScript.trainingSkillKey = Key.F6;
                    break;
                case 2:
                    GameScript.trainingSkillKey = Key.F7;
                    break;
                case 3:
                    GameScript.trainingSkillKey = Key.F8;
                    break;
#if DEBUG
                case 4:
                    GameScript.trainingSkillKey = Key.F10;
                    break;
                case 5:
                    GameScript.trainingSkillKey = Key.F11;
                    break;
#else
                case 4:
                    GameScript.trainingSkillKey = Key.F9;
                    break;
                case 5:
                    GameScript.trainingSkillKey = Key.F10;
                    break;
#endif
                case 6:
                    GameScript.trainingSkillKey = Key.F12;
                    break;
                default:
                    GameScript.trainingSkillKey = Key.F5;
                    break;
            }
        }
        public static void setPetSupportTarget(int selectedIndex)
        {
            //tab5comboAutoBuffTarget.SelectedIndex;
            // 檢查選中的索引並執行對應邏輯
            switch (selectedIndex)
            {
                case 0:
                    GameScript.petSupTarget = 0;
                    break;
                case 1:
                    GameScript.petSupTarget = 1;
                    break;
                case 2:
                    GameScript.petSupTarget = 2;
                    break;
                case 3:
                    GameScript.petSupTarget = 3;
                    break;
                case 4:
                    GameScript.petSupTarget = 4;
                    break;
                case 5:
                    GameScript.petSupTarget = 0;
                    GameScript.isPetSupportToEnemy = true;
                    break;
                default:
                    GameScript.petSupTarget = 0;
                    break;
            }
        }
        public static void setAutoBuffTarget(int selectedIndex)
        {
            // 檢查選中的索引並執行對應邏輯
            switch (selectedIndex)
            {
                case 0:
                    GameScript.buffTarget = 0;
                    break;
                case 1:
                    GameScript.buffTarget = 1;
                    break;
                case 2:
                    GameScript.buffTarget = 2;
                    break;
                case 3:
                    GameScript.buffTarget = 3;
                    break;
                case 4:
                    GameScript.buffTarget = 4;
                    break;
                default:
                    GameScript.buffTarget = 0;
                    break;
            }
        }
        public static void setMpRatioLessValue(int selectedIndex)
        {
            switch (selectedIndex)
            {
                case 0: //10
                    GameScript.mpRatioLess = 0.1;
                    break;
                case 1://20
                    GameScript.mpRatioLess = 0.2;
                    break;
                case 2://30
                    GameScript.mpRatioLess = 0.3;
                    break;
                case 3://40
                    GameScript.mpRatioLess = 0.4;
                    break;
                case 4://50
                    GameScript.mpRatioLess = 0.5;
                    break;
                case 5://60
                    GameScript.mpRatioLess = 0.6;
                    break;
                case 6://70
                    GameScript.mpRatioLess = 0.7;
                    break;
                case 7://80
                    GameScript.mpRatioLess = 0.8;
                    break;
                case 8://90
                    GameScript.mpRatioLess = 0.9;
                    break;
                default:
                    GameScript.mpRatioLess = 0.5;
                    break;
            }
        }
        public static void setAuxiliaryFunctions(int auxFunIndex)
        {
            // 檢查選中的索引並執行對應邏輯
            GameScript.usingItem = auxFunIndex;
        }
        #endregion
        #region 統整資訊
        public void checkHotkeyGetInfo()
        {
#if DEBUG
            tab2LabelAutoAttackKeyDebug.Content = "";
            tab2LabelPetSupportKeyDebug.Content = "";
            tab3LabelPetSupportKeyDebug.Content = "";
            tab4LabelPetSupportKeyDebug.Content = "";
            tab4LabelSummonAttackKeyDebug.Content = "";
            tab4LabelSummonKeyDebug.Content = ""; ;
            tab5LabelAutoBuffKeyDebug.Content = "";
            tab5LabelPetSupportKeyDebug.Content = "";
#endif
            GameScript.setGameScriptFlagDefault();
            switch (tabFunctionSelected)
            {
                case 2://tab 2 -AutoBattle
                    {
                        // 設定自動攻擊熱鍵
                        Log.Append("啟用腳本: 自動攻擊");
                        if (tab2ComboPlayerActionKey.SelectedIndex != -1)
                        {
                            int select;
                            select = tab2ComboPlayerActionKey.SelectedIndex;
                            setPlayerActionKey(select);
#if DEBUG
                            tab2LabelAutoAttackKeyDebug.Content = DebugFunction.feedBackKeyString(GameScript.playerActionKey);
#endif
                        }
                        if (tab2CheckLikeHuman.IsChecked == true)
                        {
                            Log.Append("增加隨機延遲啟用");
                            GameScript.isLikeHuman = true;
                        }

                        // 檢查寵物輔助功能
                        if (tab2CheckPetSupport.IsChecked == true)
                        {
                            Log.Append("寵物輔啟用");
                            GameScript.isPetSupport = true;
                            // 設定寵物輔助目標
                            if (tab2comboPetSupportTarget.SelectedIndex != -1)
                            {
                                int select;
                                select = tab2comboPetSupportTarget.SelectedIndex;
                                setPetSupportTarget(select);
                            }
                            // 設定寵物輔助熱鍵
                            if (tab2ComboPetActionKey.SelectedIndex != -1)
                            {
                                int select;
                                select = tab2ComboPetActionKey.SelectedIndex;
                                setPetActionKey(select);
#if DEBUG
                                tab2LabelPetSupportKeyDebug.Content = DebugFunction.feedBackKeyString(GameScript.petActionKey);
#endif
                            }
                            //
                            if (tab2CheckPetSupportJust1.IsChecked == true)
                            {
                                GameScript.isPetSupportJust1 = true;
                            }
                        }
                        else
                            GameScript.isPetSupport = false;

                        // 檢查是否自動走路
                        if (tab2CheckAutoWalk.IsChecked == true)
                        {
                            GameScript.isAutoWalk = true;
                            // 檢查是否自動走路隨機化
                            if (tab2CheckAutoWalk_OptionRandom.IsChecked == true)
                                GameScript.isAutoWalk_OptionRandom = true;
                            else
                                GameScript.isAutoWalk_OptionRandom = false;
                        }
                        else
                            GameScript.isAutoWalk = false;

                        if (tab2CheckPvpMode.IsChecked == true)
                            GameScript.isPvpMode = true;
                        else
                            GameScript.isPvpMode = false;
                    }
                    break;
                case 3://AutoDefend
                    {
                        Log.Append("啟用腳本: 自動防禦");
                        // 檢查是否寵物輔助
                        if (tab3CheckPetSupport.IsChecked == true)
                        {
                            GameScript.isPetSupport = true;
                            // 設定寵物輔助目標
                            if (tab3comboPetSupportTarget.SelectedIndex != -1)
                            {
                                int select;
                                select = tab3comboPetSupportTarget.SelectedIndex;
                                setPetSupportTarget(select);
                            }
                            // 設定寵物輔助熱鍵
                            if (tab3ComboPetActionKey.SelectedIndex != -1)
                            {
                                int select;
                                select = tab3ComboPetActionKey.SelectedIndex;
                                setPetActionKey(select);
#if DEBUG
                                tab3LabelPetSupportKeyDebug.Content = DebugFunction.feedBackKeyString(GameScript.petActionKey);
#endif
                            }
                        }
                        else
                            GameScript.isPetSupport = false;

                        // 檢查是否自動走路
                        if (tab3CheckAutoWalk.IsChecked == true)
                        {
                            GameScript.isAutoWalk = true;

                            // 檢查是否自動走路隨機化
                            if (tab3CheckAutoWalk_OptionRandom.IsChecked == true)
                                GameScript.isAutoWalk_OptionRandom = true;
                            else
                                GameScript.isAutoWalk_OptionRandom = false;
                        }
                        else
                            GameScript.isAutoWalk = false;
                    }
                    break;
                case 4://AutoEnterBattle
                    {
                        Log.Append("啟用腳本: 自動召怪");
                        // 設定自動招怪熱鍵
                        if (tab4ComboSummonerActionKey.SelectedIndex != -1)
                        {
                            int select;
                            select = tab4ComboSummonerActionKey.SelectedIndex;
                            setSummonerActionKey(select);
#if DEBUG
                            tab4LabelSummonKeyDebug.Content = DebugFunction.feedBackKeyString(GameScript.summonerActionKey);
#endif
                        }
                        // 檢查是否自召自打
                        if (tab4CheckSummonAttack.IsChecked == true)
                        {
                            GameScript.isSummonerAttack = true;
                            int select;
                            select = tab4ComboSummonerAttackKey.SelectedIndex;
                            setSummonerAttackKey(select);
#if DEBUG
                            tab4LabelSummonAttackKeyDebug.Content = DebugFunction.feedBackKeyString(GameScript.summonerAttackKey);
#endif
                        }
                        else
                            GameScript.isSummonerAttack = false;

                        // 檢查是否寵物輔助
                        if (tab4CheckPetSupport.IsChecked == true)
                        {
                            GameScript.isPetSupport = true;
                            // 設定寵物輔助目標
                            if (tab4comboPetSupportTarget.SelectedIndex != -1)
                            {
                                int select;
                                select = tab4comboPetSupportTarget.SelectedIndex;
                                setPetSupportTarget(select);
                            }
                            // 設定寵物輔助熱鍵
                            if (tab4ComboPetActionKey.SelectedIndex != -1)
                            {
                                int select;
                                select = tab4ComboPetActionKey.SelectedIndex;
                                setPetActionKey(select);
#if DEBUG
                                tab4LabelPetSupportKeyDebug.Content = DebugFunction.feedBackKeyString(GameScript.petActionKey);
#endif
                            }
                        }
                        else
                            GameScript.isPetSupport = false;
                    }
                    break;
                case 5://AutoBuff
                    {
                        Log.Append("啟用腳本: 自動增益");
                        // 設定輔助目標
                        if (tab5comboAutoBuffTarget.SelectedIndex != -1)
                        {
                            int select;
                            select = tab5comboAutoBuffTarget.SelectedIndex;
                            setAutoBuffTarget(select);
                        }
                        // 設定輔助熱鍵
                        if (tab5ComboAutoBuffKey.SelectedIndex != -1)
                        {
                            int select;
                            select = tab5ComboAutoBuffKey.SelectedIndex;
                            setPlayerActionKey(select);
#if DEBUG
                            tab5LabelAutoBuffKeyDebug.Content = DebugFunction.feedBackKeyString(GameScript.playerActionKey);
#endif
                        }
                        // 檢查是否寵物輔助
                        if (tab5CheckPetSupport.IsChecked == true)
                        {
                            GameScript.isPetSupport = true;
                            // 設定寵物輔助目標
                            if (tab5comboPetSupportTarget.SelectedIndex != -1)
                            {
                                int select;
                                select = tab5comboPetSupportTarget.SelectedIndex;
                                setPetSupportTarget(select);
                            }
                            // 設定寵物輔助熱鍵
                            if (tab5ComboPetActionKey.SelectedIndex != -1)
                            {
                                int select;
                                select = tab5ComboPetActionKey.SelectedIndex;
                                setPetActionKey(select);
#if DEBUG
                                tab5LabelPetSupportKeyDebug.Content = DebugFunction.feedBackKeyString(GameScript.petActionKey);
#endif
                            }
                        }
                        else
                            GameScript.isPetSupport = false;

                        // 檢查是否自動走路
                        if (tab5CheckAutoWalk.IsChecked == true)
                        {
                            GameScript.isAutoWalk = true;
                            // 檢查是否自動走路隨機化
                            if (tab5CheckAutoWalk_OptionRandom.IsChecked == true)
                                GameScript.isAutoWalk_OptionRandom = true;
                            else
                                GameScript.isAutoWalk_OptionRandom = false;
                        }
                        else
                            GameScript.isAutoWalk = false;
                    }
                    break;
                case 6: // 
                    {
                        Log.Append("啟用腳本: 輔助功能");
                        if (radioAutoOpenCrystallBox.IsChecked == true)
                        {
                            if (checkBagAuto.IsChecked == true)
                                GameScript.isItemRecheck = true;
                            else
                                GameScript.isItemRecheck = false;

                            GameScript.pollingItemIndex = GameScript.itemIndexDefault;
                            GameScript.pollingItemCheckTime = GameScript.itemIndexDefault;
                            GameScript.isItemEnd = false;
                            setAuxiliaryFunctions(1);
                        }
                        else if (radioAutoUsingAdjustmentPill.IsChecked == true)
                        {
                            setAuxiliaryFunctions(2);
                        }
                        else if (radioAutoNpcExchangeAlien.IsChecked == true)
                        {
                            setAuxiliaryFunctions(3);
                        }
                        else if (radioAutoBuyCrystallBox.IsChecked == true)
                        {
                            GameScript.buyCBCount = 0;
                            GameScript.buyCBCountTarget = 0;

                            int countTarget;
                            if (!int.TryParse(textBuyCBTotalCount.Text.Trim(), out countTarget))
                            {
                                countTarget = 0;
                            }
                            else
                                GameScript.buyCBCountTarget = countTarget;
                            setAuxiliaryFunctions(4);
                        }
                        else if (radioAutoLabar.IsChecked == true)
                        {
                            setAuxiliaryFunctions(5);
                        }
                        else if (radioAutoTradeItem.IsChecked == true)
                        {
                            setAuxiliaryFunctions(6);
                            GameScript.selectedItemIndex = comboTradeItemSelected.SelectedIndex;
                        }
                        else if (radioAutoBankItem.IsChecked == true)
                        {
                            setAuxiliaryFunctions(7);
                            GameScript.selectedItemIndex = comboBankItemSelected.SelectedIndex;
                            GameScript.selectedItemForm = comboClickFromSelected.SelectedIndex;
                        }
                        break;
                    }
                case 7:
                    {
                        Log.Append("啟用腳本: 自動練技");
                        int select;
                        // 檢查是否走路遇怪
                        if (tab7BattleSummon.IsChecked == true)
                        {
                            // 自動召怪
                            GameScript.trainingSkillBattleType = 1;
                            select = tab7ComboSummonerActionKey.SelectedIndex;
                            setSummonerActionKey(select);
                        }
                        else if (tab7BattleWalk.IsChecked == true)
                        {
                            // 自動走路
                            GameScript.trainingSkillBattleType = 2;
                            GameScript.isAutoWalk = true;
                            if (tab7CheckAutoWalk_OptionRandom.IsChecked == true)
                                GameScript.isAutoWalk_OptionRandom = true;
                        }
                        else
                        {
                            GameScript.trainingSkillBattleType = 0;
                        }
                        if (tab7CheckSecondRoundSkill.IsChecked == true)
                        {
                            GameScript.isSecondRoundSkill_Option = true;
                        }

                        // 檢查練技用熱鍵
                        select = tab7ComboSkillTrainingKey.SelectedIndex;
                        int roundCount;
                        if (!int.TryParse(textRoundSpellInput.Text, out roundCount))
                        {
                            roundCount = 1;
                        }
                        setTrainingSkillKey(select, roundCount);
                        if (tab7ComboSkillTrainingTarget.SelectedIndex == 5)
                        {
                            GameScript.isTrainingSkillToEnemy = true;
                        }
                        else
                        {
                            setAutoBuffTarget(select);
                        }

                        if (tab7CheckPetSupport.IsChecked == true)
                        {
                            GameScript.isPetSupport = true;
                            // 設定寵物輔助目標
                            if (tab7comboPetSupportTarget.SelectedIndex != -1)
                            {
                                int selectPK;
                                selectPK = tab7comboPetSupportTarget.SelectedIndex;
                                setPetSupportTarget(selectPK);
                            }
                            // 設定寵物輔助熱鍵
                            if (tab7ComboPetActionKey.SelectedIndex != -1)
                            {
                                int selectPK;
                                selectPK = tab7ComboPetActionKey.SelectedIndex;
                                setPetActionKey(selectPK);
#if DEBUG
                                tab7LabelPetSupportKeyDebug.Content = DebugFunction.feedBackKeyString(GameScript.petActionKey);
#endif
                            }
                        }
                        else
                            GameScript.isPetSupport = false;

                        // 檢查收尾用熱鍵
                        if (tab7AttackHotkey.IsChecked == true)
                        {
                            GameScript.trainingSkillEndType = 0;
                            select = tab7ComboSkillTrainingKey.SelectedIndex;
                            setPlayerActionKey(select);
                        }
                        else
                            GameScript.trainingSkillEndType = 1;

                        // 檢查魔力消耗條件 , 滿足則殺怪結束戰鬥
                        select = tab7ComboMPLessThan.SelectedIndex;
                        setMpRatioLessValue(select);
                    }
                    break;
                default:
                    break;
            }
        }
        #endregion
        #region 熱鍵觸發事件
        private const int WM_HOTKEY = 0x0312;
        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_HOTKEY)
            {
                // 檢查熱鍵 ID 是否符合
                if (wParam.ToInt32() == SystemSetting.HOTKEY_SCRIPT_EN)
                {
#if !DEBUG
                    //if (SystemSetting.isWin10 == true)
                    //{
                    //    Log.Append("偵測作業系統為Windows 10");
                    //    Log.Append("當前僅相容Windows 7");
                    //    Log.Append("結束腳本");
                    //    MessageBox.Show("暫不支援Win10");
                    //}
                    //else
#endif
                    {
                        if (tabFunctionSelected > 1)
                        {
                            // 初始化作弊檢查
                            GameScript.isCheckCheat = false;
                            isChangeTitle = false;

                            if (checkMouseDefault.IsChecked == true)
                                GameScript.isMouseDefault = true;
                            else
                                GameScript.isMouseDefault = false;

                            if (checkRoundDisplay.IsChecked == true)
                                GameScript.isDisplayRound = true;
                            else
                                GameScript.isDisplayRound = false;


                            // 檢查所有熱鍵.目標配置
                            checkHotkeyGetInfo();
                            // 執行相應的腳本
                            HotKeyAction_Script_EnableSwitch(this);
                        }
                        else
                        {
                            int xOffset_key = Coordinate.windowBoxLineOffset;
                            int yOffset_key = Coordinate.windowHOffset;
                            MouseSimulator.GetCurrentXY(out int currentX, out int currentY);
                            if (currentX < xOffset_key)
                                currentX = 0;
                            if (currentY < yOffset_key)
                                currentY = 0;

                            currentX = currentX - xOffset_key;
                            currentY = currentY - yOffset_key;

                            labelGetX.Content = currentX.ToString();
                            labelGetY.Content = currentY.ToString();
                        }
                    }
                    handled = true;
                }
            }
            return IntPtr.Zero;
        }
        #endregion
        #region 視窗關閉
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);

            var helper = new WindowInteropHelper(this);
            IntPtr hwnd = helper.Handle;

            SystemSetting.UnregisterHotKey(hwnd, SystemSetting.HOTKEY_SCRIPT_EN);
        }
        #endregion
        #region 測試功能按鈕
        private void btnCurrentStatusCheck_Click(object sender, RoutedEventArgs e)
        {
            SystemSetting.GetGameWindow();
            if (GameFunction.BattleCheck_Player())
            {
                MessageBox.Show(Application.Current.Resources["msgDebugGameSatsut_Player"].ToString());
            }
            else if (GameFunction.BattleCheck_Pet())
            {
                MessageBox.Show(Application.Current.Resources["msgDebugGameSatsut_Pet"].ToString());
            }
            else if (GameFunction.NormalCheck())
            {
                MessageBox.Show(Application.Current.Resources["msgDebugGameSatsut_NotBattle"].ToString());
                if (GameFunction.ItemTimeCheck())
                {
                    MessageBox.Show("當前正開啟物品欄");
                }
                else if (GameFunction.checkCheatingCheck())
                {
                    MessageBox.Show("當前有可能正在作弊檢測");
                }
            }
            else
            {
                if (GameFunction.GetBattleTimeEnd() == true)
                {
                    MessageBox.Show("回合結束 交戰中");
                    return;
                }
                MessageBox.Show(Application.Current.Resources["msgDebugGameSatsut_Exception_Status"].ToString());
            }
        }

        private void btnGetEnemyIndex_Click(object sender, RoutedEventArgs e)
        {
            SystemSetting.GetGameWindow();
            int index = GameFunction.getEnemyCoor();
            if (index > 0)
            {
                MessageBox.Show(Application.Current.Resources["msgDebugGameGetEnemyIndex"].ToString() + index + Application.Current.Resources["msgDebugGameGetEnemyUnit"].ToString());
            }
            else
            {
                MessageBox.Show(Application.Current.Resources["msgDebugGameGetEnemyError"].ToString());
            }
        }

        private void btnGetEnemyIndexBmp_Click(object sender, RoutedEventArgs e)
        {
            SystemSetting.GetGameWindow();
            DebugFunction.captureAllEnemyDotScreen();
        }

        private void btnGetTargetBmp_Click(object sender, RoutedEventArgs e)
        {
            SystemSetting.GetGameWindow();
            if (int.TryParse(this.textGetTargetBmpX.Text, out int x) &&
                int.TryParse(this.textGetTargetBmpY.Text, out int y) &&
                int.TryParse(this.textGetTargetBmpWidth.Text, out int width) &&
                int.TryParse(this.textGetTargetBmpHeight.Text, out int height))
            {
                DebugFunction.captureTargetScreen(x, y, width, height);
            }
        }

        private void btnGetAllItem_Click(object sender, RoutedEventArgs e)
        {
            SystemSetting.GetGameWindow();
            DebugFunction.captureAllItemScreen();
        }

        private void btnGetAllBmp_Click(object sender, RoutedEventArgs e)
        {
            SystemSetting.GetGameWindow();
            DebugFunction.captureAllNeedBmp();
        }

        private void btnGetBattleMp_Click(object sender, RoutedEventArgs e)
        {
            SystemSetting.GetGameWindow();
            double result = GameFunction.GetBattleMpRatio();
            double result2 = GameFunction.GetBattleHpRatio();
            MessageBox.Show($"魔力比例 '{result }', 血量比例 '{result2 }' )");
        }
        #endregion
        #region 選擇TAB時設定使用的腳本
        private void tabControlUsingMethod_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 檢查視窗是否已完全載入
            if (!isWindowLoaded)
                return;  // 如果視窗尚未載入完成，直接返回

            // 確認選中項是 TabItem
            if (e.Source is TabControl)
            {
                TabItem selectedTab = (sender as TabControl).SelectedItem as TabItem;
                if (selectedTab != null)
                {
                    // 根據選中的 Tab 執行相應的邏輯
                    switch (selectedTab.Name)
                    {
                        case "tabIndex":
                            tabFunctionSelected = 0;
                            break;
                        case "tabTestFunction":
                            tabFunctionSelected = 1;
                            break;
                        case "tabAutoBattle":
                            tabFunctionSelected = 2;
                            break;
                        case "tabAutoDefend":
                            tabFunctionSelected = 3;
                            break;
                        case "tabAutoEnterBattle":
                            tabFunctionSelected = 4;
                            break;
                        case "tabAutoBuff":
                            tabFunctionSelected = 5;
                            break;
                        case "tabAuxiliaryFunctions":
                            tabFunctionSelected = 6;
                            break;
                        case "tabSkillTraining":
                            tabFunctionSelected = 7;
                            break;
                        default:
                            tabFunctionSelected = 1;
                            break;
                    }
                }
            }
        }
        #endregion
        #region 程式語系變更
        private void LanguageChange(object sender, SelectionChangedEventArgs e)
        {
            object selected = null;
            int selectedIndex = 0;

            selected = comboLanguage.SelectedItem;
            // 檢查選項是否被選中
            if (selected != null)
            {
                // 根據 ComboBox 的選擇索引執行不同的動作
                selectedIndex = comboLanguage.SelectedIndex;

                // 檢查選中的索引並執行對應邏輯
                switch (selectedIndex)
                {
                    case 0: // 繁體中文     
                        SystemSetting.LoadResourceDictionary("zh-TW.xaml");
                        break;
                    case 1: // English   
                        SystemSetting.LoadResourceDictionary("en-US.xaml");
                        break;
                    default:
                        break;
                }
                ViewUpdate();
            }
        }
        #endregion
        #region 檢查輸入字元
        private void textGetTargetBmp_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            InputMethod.SetIsInputMethodEnabled((TextBox)sender, false);

            Regex regex = new Regex("[^0-9]+"); // 非數字的正則表達式
            e.Handled = regex.IsMatch(e.Text);  // 如果輸入非數字，則處理輸入事件為無效
        }
        #endregion
        #region 儲存/設定 欄位
        private void btnSaveSetting_Click(object sender, RoutedEventArgs e)
        {
            int battleDelay, battleDelayLong, itemDelay, sumonDelay;
            // 檢查與轉換 TextBox 的內容
            if (!int.TryParse(textKeyPressDelay.Text.Trim(), out battleDelay))
            {
                MessageBox.Show("請輸入正確的 battleDelay 數字");
                textKeyPressDelay.Focus();
                return;
            }

            if (!int.TryParse(textMousePressDelay.Text.Trim(), out battleDelayLong))
            {
                MessageBox.Show("請輸入正確的 battleDelayLong 數字");
                textMousePressDelay.Focus();
                return;
            }

            if (!int.TryParse(textSupportDelay.Text.Trim(), out itemDelay))
            {
                MessageBox.Show("請輸入正確的 itemDelay 數字");
                textSupportDelay.Focus();
                return;
            }

            if (!int.TryParse(textSummonDelay.Text.Trim(), out sumonDelay))
            {
                MessageBox.Show("請輸入正確的 招怪延遲 數字");
                textSummonDelay.Focus();
                return;
            }

            Properties.Settings.Default.itemDelay = itemDelay;
            Properties.Settings.Default.battleDelay = battleDelay;
            Properties.Settings.Default.battleDelayLong = battleDelayLong;
            Properties.Settings.Default.summonDelay = sumonDelay;
            Properties.Settings.Default.Save();

            GameScript.openItemDelay = Properties.Settings.Default.itemDelay;
            GameScript.keyPressDelay = Properties.Settings.Default.battleDelay;
            GameScript.keyPressDelayLong = Properties.Settings.Default.battleDelayLong;
            GameScript.summonDelay = Properties.Settings.Default.summonDelay;

            textSupportDelay.Text = GameScript.openItemDelay.ToString();
            textKeyPressDelay.Text = GameScript.keyPressDelay.ToString();
            textMousePressDelay.Text = GameScript.keyPressDelayLong.ToString();
            textSummonDelay.Text = GameScript.summonDelay.ToString();
            MessageBox.Show("儲存完成");
        }

        private void btnDefaultSetting_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.itemDelay = 150;
            Properties.Settings.Default.battleDelay = 50;
            Properties.Settings.Default.battleDelayLong = 100;
            Properties.Settings.Default.summonDelay = 0;
            Properties.Settings.Default.Save();

            GameScript.openItemDelay = Properties.Settings.Default.itemDelay;
            GameScript.keyPressDelay = Properties.Settings.Default.battleDelay;
            GameScript.keyPressDelayLong = Properties.Settings.Default.battleDelayLong;
            GameScript.summonDelay = Properties.Settings.Default.summonDelay;

            textSupportDelay.Text = GameScript.openItemDelay.ToString();
            textKeyPressDelay.Text = GameScript.keyPressDelay.ToString();
            textMousePressDelay.Text = GameScript.keyPressDelayLong.ToString();
            textSummonDelay.Text = GameScript.summonDelay.ToString();
            MessageBox.Show("初始化完成");
        }
        #endregion
    }
}
