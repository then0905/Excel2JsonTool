using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using Newtonsoft.Json;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using System.Runtime.InteropServices;
using System.ComponentModel;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace Excel2JsonTool
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        XmlDocument doc = new XmlDocument();

        // 紀錄最後的導入Excel路徑
        string lastExcelPath = "";
        // 紀錄最後的輸出Json路徑
        string lastJsonPath = "";

        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 轉換按鈕功能 並且生成json檔案
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            //Json導出路徑
            string jsonPath = JsonInputTextBox.Text;
            //Excel列表
            string excelPath = ExcelInputTextBox.Text;

            //檢查導入輸出路徑是否正確
            if (string.IsNullOrEmpty(jsonPath) || string.IsNullOrEmpty(excelPath))
            {
                WaitingText.Visibility = Visibility.Visible;
                WaitingText.Content = "錯誤：未正確選擇導入或輸出路徑";
                return;
            }

            //關閉按鈕 開啟等待中 文字
            ConvertButton.IsEnabled = false;
            WaitingText.Visibility = Visibility.Visible;
            WaitingText.Content = "尋找檔案中...";
            List<string> excelPathList = excelPath.Split('\n').ToList();

            //全部Json
            List<string> JsonList = new List<string>();
            //完成的Json
            List<string> FinishJsonList = new List<string>();
            await Task.Run(() =>
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // 設置EPPlus的許可証內容

                //取出所有的excel
                foreach (string excelPathItem in excelPathList)
                {

                    //檢查空值
                    if (string.IsNullOrEmpty(excelPathItem)) continue;

                    using (var package = new ExcelPackage(new FileInfo(excelPathItem)))
                    {
                        foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
                        {
                            if (worksheet.Name.Contains(".json") && !worksheet.Name.Contains("#"))
                            {
                                string sheetName = worksheet.Name;
                                JsonList.Add(sheetName);
                            }
                        }
                    }
                }
                //設定工作狀態
                UpdateWorking(JsonList);
            });

            WaitingText.Content = "正在轉檔中...";

            //暫存全部的轉檔資料
            var tempJsonData = new Dictionary<string, List<Dictionary<string, object>>>();

            await Task.Run(() =>
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;// 關閉新許可模式通知

                foreach (string excelPathItem in excelPathList)
                {
                    if (string.IsNullOrEmpty(excelPathItem)) continue;
                    using (var package = new ExcelPackage(new FileInfo(excelPathItem)))
                    {
                        foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
                        {
                            // 找出工作表内有.json的工作表
                            if (worksheet.Name.Contains(".json"))
                            {
                                //暫存含有#字表單的標題
                                string jsonListTitle = worksheet.Cells[1, 2].Value.ToString();
                                //找出內容最多到第幾列
                                int lastColumnNumber = worksheet.Dimension.End.Column;

                                //找出內容最多到第幾行
                                int lastRowNumber = worksheet.Dimension.End.Row;

                                // 創建空的字典存儲資料 (以Json表格為單位)
                                var data = new List<Dictionary<string, object>>();

                                // 從2開始，跳過標題行
                                for (int rowNumber = 2; rowNumber <= lastRowNumber; rowNumber++)
                                {
                                    //創建空的字典儲存資料(以行列為單位)
                                    var dataRow = new Dictionary<string, object>();
                                    // 以列為單位
                                    for (int columnNumber = 1; columnNumber <= lastColumnNumber; columnNumber++)
                                    {
                                        // 該格為空值或是該標題列沒資料，跳過
                                        if (worksheet.Cells[rowNumber, columnNumber].Value == null ||
                                            string.IsNullOrEmpty(worksheet.Cells[rowNumber, columnNumber].Value.ToString()) ||
                                            worksheet.Cells[1, columnNumber].Value == null)
                                            continue;

                                        // 參數名稱，讀取標題列的值
                                        string headerName = worksheet.Cells[1, columnNumber].Value.ToString();
                                        // 參數值
                                        object cellValue = worksheet.Cells[rowNumber, columnNumber].Value;

                                        if (cellValue is string && ((string)cellValue).Contains(","))
                                        {
                                            // 若參數包含':' 分割成List
                                            cellValue = ((string)cellValue).Split(',').ToList();
                                        }
                                        //若參數值可以為數字 判斷整數或是小數點
                                        if (cellValue is double)
                                        {
                                            double doubleValue = (double)cellValue;
                                            if (doubleValue == (int)doubleValue)
                                            {
                                                cellValue = (int)doubleValue;
                                            }
                                        }

                                        //寫入此次
                                        dataRow.Add(headerName, cellValue);
                                    }
                                    if (dataRow.Count > 0)
                                        data.Add(dataRow);
                                }

                                var jsonSettings = new JsonSerializerSettings
                                {
                                    Formatting = Newtonsoft.Json.Formatting.Indented,
                                };

                                string sheetName = worksheet.Name;

                                //若表單含有#字
                                if (sheetName.Contains('#'))
                                {
                                    //取得該表單正確的json名稱
                                    string mainSheetName = sheetName.Split('#')[0];

                                    // 如果主json文件還沒在 tempJsonData 中，創造一個
                                    if (!tempJsonData.ContainsKey(mainSheetName))
                                        tempJsonData[mainSheetName] = new List<Dictionary<string, object>>();

                                    // 尋找主表裡已有的和現在這個子表第一個key相同的值
                                    foreach (Dictionary<string, object> mainItem in tempJsonData[mainSheetName])
                                    {
                                        var linkKey = mainItem.First().Value;  // 主表的第一個value
                                        foreach (var item in data)
                                        {
                                            var subLinkKey = item.First().Value;  //子表的第一個value
                                            if (linkKey == subLinkKey)
                                            {
                                                // 若找到，創造一個新的子表列表
                                                if (!mainItem.ContainsKey(jsonListTitle))
                                                {
                                                    mainItem[jsonListTitle] = new List<Dictionary<string, object>>();
                                                }
                                         (mainItem[jsonListTitle] as List<Dictionary<string, object>>).Add(item);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    // 如果工作表名稱不含 '#', 當作主json文件處理
                                    tempJsonData[sheetName] = data;
                                }
                            }
                        }
                    }
                }
            });

            await Task.Run(() =>
            {
                // 最後把 tempJsonData 寫入 JSON 檔案
                foreach (var item in tempJsonData)
                {
                    var jsonSettings = new JsonSerializerSettings
                    {
                        Formatting = Newtonsoft.Json.Formatting.Indented,
                    };
                    string json = JsonConvert.SerializeObject(item.Value, jsonSettings);
                    // 創建json檔案
                    File.WriteAllText($"{jsonPath}\\{item.Key}", json);
                    FinishJsonList.Add(item.Key);
                    // 設定工作狀態
                    UpdateWorking(JsonList, FinishJsonList);
                }
            });

            ConvertButton.IsEnabled = true;
            WaitingText.Content = "轉檔完成";
        }

        /// <summary>
        /// 導入excel路徑
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BrowseExcel_Click(object sender, RoutedEventArgs e)
        {
            // 建立一個 CommonOpenFileDialog 來讓使用者選擇一個資料夾
            var folderDialog = new Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog();
            lastExcelPath = Properties.Settings.Default.LastSelectedExcelFolder;
            //若有路徑紀錄 使用該紀錄
            if (!string.IsNullOrEmpty(lastExcelPath))
            {
                folderDialog.InitialDirectory = lastExcelPath;
            }
            //設定為選擇資料夾模式
            folderDialog.IsFolderPicker = true;

            if (folderDialog.ShowDialog() == Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok)
            {
                //獲取選擇的資料夾路徑
                string folderPath = folderDialog.FileName;

                //取出所有excel檔案 並過濾~$的暫存檔案
                var excelFiles = Directory.EnumerateFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly)
                    .Where(file => (file.ToLower().EndsWith("xls")
                        || file.ToLower().EndsWith("xlsx")
                        || file.ToLower().EndsWith("xlsm")) && !file.Contains("~$"));

                //重置textbox內容
                ExcelInputTextBox.Text = string.Empty;

                foreach (var excelFile in excelFiles)
                {
                    ExcelInputTextBox.Text += excelFile + "\n"; //將找到的檔案路徑顯示在TextBox中
                }

                //刷新路徑紀錄
                Properties.Settings.Default.LastSelectedExcelFolder = folderDialog.FileName;
                Properties.Settings.Default.Save();
            }
        }

        /// <summary>
        /// 輸出Json路徑 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExportJsonPath_Click(object sender, RoutedEventArgs e)
        {
            // 建立一個FolderBrowserDialog來讓使用者選擇一個資料夾
            //var folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            //if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //{
            //    // 紀錄選擇的資料夾路徑
            //    JsonInputTextBox.Text = folderBrowserDialog.SelectedPath;
            //}
            // 建立一個 CommonOpenFileDialog 來讓使用者選擇一個資料夾
            var folderDialog = new Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog();
            lastJsonPath = Properties.Settings.Default.LastSelectedJsonFolder;
            if (!string.IsNullOrEmpty(lastJsonPath))
            {
                folderDialog.InitialDirectory = lastJsonPath;
            }
            //設定為選擇資料夾模式
            folderDialog.IsFolderPicker = true;

            if (folderDialog.ShowDialog() == Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok)
            {
                JsonInputTextBox.Text = folderDialog.FileName;
                //刷新路徑紀錄
                Properties.Settings.Default.LastSelectedJsonFolder = folderDialog.FileName;
                Properties.Settings.Default.Save();
            }
        }

        /// <summary>
        /// 更新工作狀態
        /// </summary>
        private async void UpdateWorking(List<string> worksheets, List<string> finishWorksheets = null)
        {

            await System.Windows.Application.Current.Dispatcher.Invoke(async () =>
            {
                ResultTextBox.Content = string.Empty;
                //if (finishWorksheets != null && finishWorksheets.Count > 0)
                //    ResultTextBox.Content = "已完成(" + finishWorksheets.Count.ToString() + "/" + worksheets.Count.ToString() + ")";
                //else
                //    ResultTextBox.Content = "已完成(0" + "/" + worksheets.Count.ToString() + ")";
                foreach (string worksheet in worksheets)
                {
                    if (finishWorksheets != null && finishWorksheets.Count > 0)
                    {

                        if (finishWorksheets.Any(x => x.Contains(worksheet)))
                            ResultTextBox.Content += worksheet + ":已完成" + "\n";
                        else
                            ResultTextBox.Content += worksheet + ":未完成" + "\n";
                    }
                    else
                        ResultTextBox.Content += worksheet + ":未完成" + "\n";
                }
                await Task.Delay(10);
                //await Task.Delay(100);
            });

        }
    }
}
