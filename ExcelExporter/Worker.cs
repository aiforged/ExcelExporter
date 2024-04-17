namespace ExcelExporter
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private readonly Models.Config _config;
        private ExcelGenerationWorker _genWorker;

        public Worker(ILogger<Worker> logger, Models.Config config)
        {
            _logger = logger;
            _config = config;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _genWorker = new ExcelGenerationWorker(_logger, _config, stoppingToken);

            while (!stoppingToken.IsCancellationRequested)
            {
                if (_logger.IsEnabled(LogLevel.Information))
                {
                    _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                }

                await _genWorker.CheckAndGenerateExcelAsync();

                await Task.Delay(30000, stoppingToken);
            }
        }
    }
}
