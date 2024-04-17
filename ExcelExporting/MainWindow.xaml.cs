using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Storage.Pickers;
using Windows.Storage.Streams;
using Windows.Storage;
using Windows.UI.Popups;
using Syncfusion.UI.Xaml.DataGrid.Export;
using Syncfusion.XlsIO;
using Syncfusion.UI.Xaml.DataGrid;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace ExcelExporting
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        public MainWindow()
        {
            this.InitializeComponent();
        }

        private void OnExportToExcelClick(object sender, RoutedEventArgs e)
        {
            var options = new DataGridExcelExportOptions();
            options.CanExportStackedHeaders = ExportStackedHeaders.IsChecked == true;
            options.ExcelVersion = ExcelVersion.Excel2013;
            if (ColumnStyle.IsChecked == true)
                options.CellsExportHandler = CellsExportHandler;

            if (OrderIDColumn.IsChecked == false)
                options.ExcludedColumns.Add("OrderID");

            if (OrderDateColumn.IsChecked == false)
                options.ExcludedColumns.Add("OrderDate");

            if (ShippingCityColumn.IsChecked == false)
                options.ExcludedColumns.Add("ShipCity");

            if (ShippingCountryColumn.IsChecked == false)
                options.ExcludedColumns.Add("ShipAddress");

            if (QuantityColumn.IsChecked == false)
                options.ExcludedColumns.Add("Quantity");

            if (UnitPriceColumn.IsChecked == false)
                options.ExcludedColumns.Add("UnitPrice");

            var excelEngine = sfDataGrid.ExportToExcel(sfDataGrid.View, options);
            var workBook = excelEngine.Excel.Workbooks[0];
            MemoryStream outputStream = new MemoryStream();
            workBook.SaveAs(outputStream);
            SaveExcelWorkbook(outputStream, "OrderDetails");
        }

        private void CellsExportHandler(object sender, DataGridCellExcelExportOptions e)
        {
            if (e.ColumnName == "UnitPrice")
            {
                e.Range.CellStyle.ColorIndex = ExcelKnownColors.Blue_grey;
                e.Range.CellStyle.Font.Color = ExcelKnownColors.Light_yellow;
            }
        }


        async void SaveExcelWorkbook(MemoryStream stream, string filename)
        {
            StorageFile stFile;
            if (!(Windows.Foundation.Metadata.ApiInformation.IsTypePresent("Windows.Phone.UI.Input.HardwareButtons")))
            {
                FileSavePicker savePicker = new FileSavePicker();
                savePicker.DefaultFileExtension = ".xlsx";
                savePicker.SuggestedFileName = filename;
                savePicker.FileTypeChoices.Add("Excel Documents", new List<string>() { ".xlsx" });
                var hwnd = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                WinRT.Interop.InitializeWithWindow.Initialize(savePicker, hwnd);
                stFile = await savePicker.PickSaveFileAsync();
            }
            else
            {
                StorageFolder local = Windows.Storage.ApplicationData.Current.LocalFolder;
                stFile = await local.CreateFileAsync(filename, CreationCollisionOption.ReplaceExisting);
            }

            if (stFile != null)
            {
                using (IRandomAccessStream zipStream = await stFile.OpenAsync(FileAccessMode.ReadWrite))
                {
                    //Write compressed data from memory to file
                    using (Stream outstream = zipStream.AsStreamForWrite())
                    {
                        byte[] buffer = stream.ToArray();
                        outstream.Write(buffer, 0, buffer.Length);
                        outstream.Flush();
                    }
                }
                //Launch the saved Excel file
                await Windows.System.Launcher.LaunchFileAsync(stFile);
            }
        }

        private void OnDataGridSelectionChanged(object sender, Syncfusion.UI.Xaml.Grids.GridSelectionChangedEventArgs e)
        {
            if (this.sfDataGrid.SelectedItems.Count > 0)
            {
                this.exportSelectedItems.IsEnabled = true;
                NoteTextBlock.Visibility = Microsoft.UI.Xaml.Visibility.Collapsed;
            }
            else
            {
                this.exportSelectedItems.IsEnabled = false;
                NoteTextBlock.Visibility = Microsoft.UI.Xaml.Visibility.Visible;
            }
        }

        private void OnExportSelectedRowsClick(object sender, RoutedEventArgs e)
        {
            var options = new DataGridExcelExportOptions();
            options.ExcelVersion = ExcelVersion.Excel2013;
            if (rowStyleCustomizationCheckBox.IsChecked == true)
                options.GridExportHandler = GridExportHandler;

            var excelEngine = sfDataGrid.ExportToExcel(sfDataGrid.SelectedItems, options);
            var workBook = excelEngine.Excel.Workbooks[0];
            MemoryStream outputStream = new MemoryStream();
            workBook.SaveAs(outputStream);
            SaveExcelWorkbook(outputStream, "SelectedOrders");
        }

        private void GridExportHandler(object sender, DataGridExcelExportStartOptions e)
        {
            if (e.CellType == ExportCellType.RecordCell)
            {
                e.Style.ColorIndex = ExcelKnownColors.Sea_green;
                e.Style.Font.Color = ExcelKnownColors.White;
            }
        }
    }
}