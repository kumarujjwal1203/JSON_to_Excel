using System;
using System.Windows;
using System.Windows.Threading;
using GSTJsonToExcel.Services.Implementations;
using GSTJsonToExcel.ViewModels;

namespace GSTJsonToExcel
{
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            var logger = new LoggingService();

            DispatcherUnhandledException += (s, args) =>
            {
                string detail = args.Exception.ToString();
                System.IO.File.WriteAllText(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "GSTJsonToExcel_Crash.log"), detail);
                logger.LogError("Unhandled UI Exception: " + detail, args.Exception);
                MessageBox.Show(
                    $"An unexpected UI error occurred:\n{args.Exception.InnerException?.Message ?? args.Exception.Message}",
                    "GST JSON to Excel Converter",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                args.Handled = true;
            };

            try
            {
                var fileService = new FileService();
                var parserService = new JsonParserService(logger);
                var exportService = new ExcelExportService(logger);
                var integrityService = new IntegrityCheckService(logger);
                var octaBuilderService = new OctaGstBuilderService(logger);

                var classifierService = new GstClassifierService(logger);
                var duplicateDetectorService = new DuplicateDetectorService(logger);
                var scannerService = new FileScannerService(classifierService, duplicateDetectorService, fileService, logger);
                var batchService = new BatchConversionService(parserService, exportService, integrityService, logger, octaBuilderService);

                var viewModel = new MainViewModel(
                    scannerService,
                    batchService,
                    fileService,
                    logger);

                var mainWindow = new MainWindow(viewModel);
                mainWindow.Show();
            }
            catch (Exception ex)
            {
                string detail = ex.ToString();
                System.IO.File.WriteAllText(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "GSTJsonToExcel_Crash.log"), detail);
                logger.LogError("Startup Exception: " + detail, ex);
                throw;
            }
        }
    }
}
