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
            Bitmap screenshot_NormalCheckPoint = BitmapFunction.CaptureScreen(x_LU, y_LU, 9, 30);

            // 比對圖像
            double Final_NCP = BitmapFunction.CompareImages(screenshot_NormalCheckPoint, NormalCheckPoint);
            Log.Append($"一般狀態檢測值 '{Final_NCP}')\n");
            Log.Append($"=============================");

            if (Final_NCP > 80)
                return true;
            else
                return false;
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
            Bitmap screenshot_keyBar = BitmapFunction.CaptureScreen(x_key, y_key, 33, 34);

            // 比對圖像
            double Final_KeyBar = BitmapFunction.CompareImages(screenshot_keyBar, fight_keybarBMP);
            Log.Append($"玩家檢測值為 '{Final_KeyBar}')\n");
            Log.Append($"=============================");

            if (Final_KeyBar > 80)
                return true;
            else
                return false;
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
            Bitmap screenshot_keyBarPet = BitmapFunction.CaptureScreen(x_key, y_key, 33, 34);

            // 比對圖像
            double Final_KeyBar = BitmapFunction.CompareImages(screenshot_keyBarPet, fight_keybarPetBMP);
            Log.Append($"寵物檢測值為 '{Final_KeyBar}')\n");
            Log.Append($"=============================");

            if (Final_KeyBar > 80)
                return true;
            else
                return false;
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
            Bitmap screenshot_itemTimeBitmap = BitmapFunction.CaptureScreen(x_ItemCP, y_ItemCP, 20, 20);

            // 比對圖像
            double Final_ICP = BitmapFunction.CompareImages(screenshot_itemTimeBitmap, itemTimeBitmap);


            Log.Append($"物品欄檢測值為 '{Final_ICP}')\n");


            if (Final_ICP > 50)
                return true;
            else
                return false;
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
            Bitmap screenshot_ItemCBBmpGet;
            System.Drawing.Color PillExistColor = System.Drawing.Color.FromArgb(116, 85, 136);
            // 取得所有物品欄位                   
            GameFunction.GetItemCoor(input, out xo, out yo);

            screenshot_ItemCBBmpGet = BitmapFunction.CaptureScreen(xo + xOffset, yo + yOffset, 1, 1);
            // 比對圖像
            double Final_ICP = BitmapFunction.CalculateColorRatio(screenshot_ItemCBBmpGet, PillExistColor);


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
            Bitmap screenshot_ItemCBBmpGet;
            System.Drawing.Color PillExistColor = System.Drawing.Color.FromArgb(113, 148, 156);
            // 取得所有物品欄位                   
            GameFunction.GetItemCoor(input, out xo, out yo);
            screenshot_ItemCBBmpGet = BitmapFunction.CaptureScreen(xo + xOffset, yo + yOffset, 1, 1);

            // 比對圖像
            double Final_ICP = BitmapFunction.CalculateColorRatio(screenshot_ItemCBBmpGet, PillExistColor);


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
            Bitmap screenshot_AP_UsingOK = BitmapFunction.CaptureScreen(x_key, y_key, 10, 10);

            // 比對圖像
            double Final_KeyBar = BitmapFunction.CompareImages(screenshot_AP_UsingOK, AP_UsingOK_BMP);


            Log.Append($"CheckItemAP_Yes 檢測值為 '{Final_KeyBar}')\n");


            if (Final_KeyBar > 80)
                return true;
            else
                return false;
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
            Bitmap screenshot_Mp = BitmapFunction.CaptureScreen(x, y, 90, 1);
            result = BitmapFunction.CalculateColorRatio(screenshot_Mp, MpExistColor);
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
            Bitmap screenshot_Hp = BitmapFunction.CaptureScreen(x, y, 90, 1);
            result = BitmapFunction.CalculateColorRatio(screenshot_Hp, HpExistColor);
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
            Bitmap screenshot_cheatCheckBitmap = BitmapFunction.CaptureScreen(x_cheatCheck, y_cheatCheck, 40, 10);

            double Final_CCC = BitmapFunction.CompareImages(screenshot_cheatCheckBitmap, cheatCheckBitmap);
            Log.Append($"作弊檢測值為 '{Final_CCC}')\n");

            if (Final_CCC > 90)
                return true;
            else
                return false;
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
            Bitmap screenshot_Hp = BitmapFunction.CaptureScreen(x, y, 160, 1);
            result = BitmapFunction.CalculateColorRatio(screenshot_Hp, BTExistColor);

            if (result == 0)
                return true;
            else
                return false;
        }
    }

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
                    //MessageBox.Show($"視窗 '{windowTitle}' 的左邊頂點座標為：({GameWindowsInfor.Left}, {GameWindowsInfor.Top})\n" +
                    //$"右邊底部座標為：({GameWindowsInfor.Right}, {GameWindowsInfor.Bottom})\n" +
                    //$"視窗長寬：({GameWindowsInfor.Right - GameWindowsInfor.Left}, {GameWindowsInfor.Bottom - GameWindowsInfor.Top})\n");
                    return true;
                }
                else
                {
                    //MessageBox.Show($"無法獲取視窗 '{windowTitle}' 的座標");
                    return false;
                }
            }
            else
            {
                //MessageBox.Show($"找不到名稱為 '{windowTitle}' 的視窗");
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

