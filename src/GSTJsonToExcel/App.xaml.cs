using System.Windows;
using GSTJsonToExcel.Services.Implementations;
using GSTJsonToExcel.ViewModels;

namespace GSTJsonToExcel
{
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            var logger = new LoggingService();
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
    }
}
