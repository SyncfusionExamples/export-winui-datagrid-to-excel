# Export To Excel in WinUI DataGrid

The [WinUI DataGrid](https://help.syncfusion.com/winui/datagrid/overview) supports exporting data to excel. Export unbound rows, unbound columns, merged cells, stacked headers, and Details View while exporting.

The following assemblies needs to be added for exporting to excel.

* Syncfusion.GridExport.WinUI
* Syncfusion.XlsIO.NET

For NuGet package, install the [Syncfusion.GridExport.WinUI](https://www.nuget.org/packages/Syncfusion.GridExport.WinUI) package.

Export the SfDataGrid to excel by using the [ExportToExcel](https://help.syncfusion.com/cr/winui/Syncfusion.UI.Xaml.DataGrid.Export.DataGridExcelExportExtensions.html#Syncfusion_UI_Xaml_DataGrid_Export_DataGridExcelExportExtensions_ExportToExcel_Syncfusion_UI_Xaml_DataGrid_SfDataGrid_Syncfusion_UI_Xaml_Data_ICollectionViewAdv_Syncfusion_UI_Xaml_DataGrid_Export_DataGridExcelExportOptions_) extension method in the [Syncfusion.UI.Xaml.DataGrid.Export](https://help.syncfusion.com/cr/winui/Syncfusion.UI.Xaml.DataGrid.Export.html) namespace.

**C#**
```csharp
using Syncfusion.UI.Xaml.DataGrid.Export;
var options = new DataGridExcelExportOptions();
var excelEngine = dataGrid.ExportToExcel(dataGrid.View, options);
var workBook = excelEngine.Excel.Workbooks[0];
MemoryStream stream = new MemoryStream();
workBook.SaveAs(stream);
Save(stream, "Sample");

async void Save(MemoryStream stream, string filename)
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
            //Write the compressed data from the memory to the file
            using (Stream outstream = zipStream.AsStreamForWrite())
            {
                byte[] buffer = stream.ToArray();
                outstream.Write(buffer, 0, buffer.Length);
                outstream.Flush();
            }
        }
        //Launch the saved Excel file.
        await Windows.System.Launcher.LaunchFileAsync(stFile);
    }
}
```

Take a moment to peruse the [WinUI SfDataGrid - Excel Export Documentation](https://help.syncfusion.com/winui/datagrid/export-to-excel), where you can find more about excel exporting with code examples.