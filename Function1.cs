using System;
using System.Collections.Generic;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;

namespace paymentreminder
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static void Run([TimerTrigger("0 */1 * * * *")]TimerInfo myTimer, ILogger log)
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");

            string file = "payments.xlsx";
            string worksheet = "PAyments";
            string table = "Payments";

            ExcelHelper excelHelper = new ExcelHelper();

            List<Expense> expenses = excelHelper.RetrieveData(file, worksheet, table, log);
            log.LogInformation(string.Format("No of elements in the list {0}", expenses.Count));
            excelHelper.GenerateEmail(expenses, log);
        }
    }
}
