﻿using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace paymentreminder
{
    class ExcelHelper
    {
        string accessToken = Environment.GetEnvironmentVariable("AccessToken");
        string baseurl = "https://graph.microsoft.com/v1.0/";

        public void GenerateEmail (List<Expense> expenses, ILogger log)
        {
            string mailaddress = "hansamaligamage@gmail.com";
            string subject = "Your expenses for May";
            StringBuilder items = new StringBuilder();
            double totalAmount = 0;
            
            foreach(Expense expense in expenses)
            {
                items.Append(expense.Description + " , ");
                totalAmount += expense.Price;
            }
            string content = string.Format("You have to pay {0}$ for {1}", totalAmount, items.ToString());
            log.LogInformation(string.Format("Email Content {0}", content));

            SmtpClient smtpClient = EmailClientBuilder();

            Email email = new Email { MailRecipientsTo = mailaddress, Content = content, Subject = subject };

            var emailMessage = MessageBuilder(email);
            try
            {
                smtpClient.Send(emailMessage);
                log.LogInformation("Email sent");
            }
            catch (Exception ex)
            {
                log.LogInformation(ex.Message);
            }
        } 

        private static SmtpClient EmailClientBuilder()
        {
            string emailServer = Environment.GetEnvironmentVariable("emailServer");
            int emailPort = Convert.ToInt32(Environment.GetEnvironmentVariable("emailPort"));
            string emailCredentialUserName = Environment.GetEnvironmentVariable("emailCredentialUserName");
            string emailCredentialPassword = Environment.GetEnvironmentVariable("emailCredentialPassword");
            SmtpClient smtpClient = new SmtpClient(emailServer, emailPort);
            smtpClient.Credentials = new System.Net.NetworkCredential(emailCredentialUserName, emailCredentialPassword);
            return smtpClient;
        }

        private static MailMessage MessageBuilder(Email email)
        {
            string fromAddress = Environment.GetEnvironmentVariable("fromAddress");
            
            MailMessage mailMsg = new MailMessage();

            mailMsg.To.Add(new MailAddress(email.MailRecipientsTo));
            mailMsg.From = new MailAddress(fromAddress);
            mailMsg.Subject = email.Subject;
            mailMsg.Body = email.Content;
            return mailMsg;
        }

        public List<Expense> RetrieveData(string file, string worksheet, string table, ILogger log)
        {
            string rowcontent = string.Empty;
            List<Expense> expenses = new List<Expense>();
            List<Expense> filteredExpenses = new List<Expense>();

            var client = new RestClient(string.Format("{0}me/drive/root/children/{1}/workbook/worksheets/{2}/Tables/{3}/rows", baseurl, file, worksheet, table));
            log.LogInformation("REST client is created");
            var request = new RestRequest(Method.GET);
            request.AddHeader("Authorization", accessToken);

            IRestResponse response = client.Execute(request);
            if (response.IsSuccessful)
            {
                log.LogInformation(string.Format("File found {0}", response.IsSuccessful));
                string content = response.Content;
                JObject filedetails = (JObject)JsonConvert.DeserializeObject(content);
                rowcontent = filedetails["value"].ToString();
                var rowdata = JsonConvert.DeserializeObject<List<TableRequest>>(rowcontent);
                for(int i = 0; i < rowdata.Count; i++)
                {
                    var data = rowdata[i];
                    var row = data.values[0];
                    Expense expense = new Expense { Description = row[1], Price = Convert.ToDouble(row[2]), DueDate = string.IsNullOrEmpty(row[3]) ? null : (DateTime?)DateTime.FromOADate(Convert.ToDouble(row[3])) };
                    expenses.Add(expense);
                }

                DateTime fromDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                DateTime toDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month));

                filteredExpenses = expenses.Where(p => p.DueDate >= fromDate && p.DueDate < toDate).ToList();

            }
            return filteredExpenses;
        }

    }
}
