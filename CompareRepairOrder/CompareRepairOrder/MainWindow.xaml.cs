using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows;
using System.Linq;
using CompareExcelItem.Service;
using CompareRepairOrder.Model;
using CompareRepairOrder.Service;
using System.Security.Policy;

namespace CompareRepairOrder
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnSelectExcel(object sender, RoutedEventArgs e)
        {
            OpenFileDialog chrooseFileDialog = new OpenFileDialog();
            chrooseFileDialog.DefaultExt = ".xlsx";
            chrooseFileDialog.Filter = "Excel files(.xlsx;)|*.xlsx;";
            chrooseFileDialog.Multiselect = false;
            chrooseFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            Nullable<bool> selected = chrooseFileDialog.ShowDialog();
            string defaultSaveExcelPath = string.Empty;
            string excleFilePath = "";
            
            try
            {
                if (selected == true)
                {
                    excleFilePath = chrooseFileDialog.FileName;
                    FileInfo filePath = new FileInfo(excleFilePath);
                    ExcelPackage ep = new ExcelPackage(filePath);

                    List<string> viewSheets = new List<string>();
                    foreach (var item in ep.Workbook.Worksheets)
                    {
                        viewSheets.Add(item.ToString());
                    }

                    if (viewSheets.Count > 0)
                        CompareExcelData(viewSheets, excleFilePath);

                }
                else
                    MessageBox.Show("Please choose excel file", "Info");

            }
            catch (Exception ex)
            {
                throw;
            }
        }


        /// <summary>
        /// 使用調整單的品號去，找退料檔裡面出現幾次，並產生免安裝執行檔
        /// </summary>
        private void CompareExcelData(List<string> excelSheets, string excleFilePath)
        {
            ImportFileService importFileService = new ImportFileService();
            ExportFileService exportFileService = new ExportFileService();
            CommonService commonService = new CommonService();

            try
            {
                //維修單資料
                List<RepairOrder> originalData = new List<RepairOrder>();
                string newPartName = txtNewPartName.Text;
                string repairOrderName = txtRepairOrderName.Text;
                int excelColumns = 0;
                int.TryParse(txtExcelColumns.Text, out excelColumns);
                originalData = importFileService.ExcelToList(excleFilePath, excelColumns, newPartName, repairOrderName);

                //Ref https://ithelp.ithome.com.tw/articles/10196181
                //找出件號在維修單號出現幾次，並列舉所有維修單號
                List<RepairOrder> newPartList = originalData.GroupBy(x => x.NewPartNumber, x => x.RepairOrderNumber, (partNumber, orderNumber) => new RepairOrder
                {
                    NewPartNumber = partNumber,
                    Quantity = orderNumber.Count(),
                    RepairOrderNumberList = orderNumber.ToList()

                }).ToList();


                DataTable dtResult = commonService.RepairOrderListToDataTable(newPartList, newPartName, repairOrderName);

                ResponseMessage response = exportFileService.DataTableToExcelFile(dtResult, excleFilePath);
                if (response.Success)
                {
                    MessageBox.Show("比對完成，已匯出到選擇Excel的路徑下", "Info");
                }
                else
                {
                    MessageBox.Show(response.ErrorMsg, "error");
                }

                //全自動轉型 弱型別轉強型別
                //var result = DataTableExtensions.ToList<RepairOrder>(dtRepairOrder).ToList();
                //List<RepairOrder> returnList = result as List<RepairOrder

                //Compare excel data
                // DataTable dt = importFileService.CompareRevisionAndReturn(dtRevision, returnList);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "error");
            }

        }

    }
}
