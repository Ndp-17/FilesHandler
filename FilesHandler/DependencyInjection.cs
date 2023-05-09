using FilesHandler.Interfaces;
using FilesHandler.Services;
using Microsoft.Extensions.DependencyInjection;

namespace FilesHandler
{
    public static class DependencyInjection
    {
        public static void AddExcelHandler( this IServiceCollection services) 
        {
            services.AddSingleton<IExcelHandler, ExcelHandler>();
        }
    }
}