# Azure function to read csv file and send email
This is a azure function timer trigger written in c# - .NET Core 3.1 to read a csv file using Microsoft Grapg API and send an email using Send Grid

# Technology stack
* .NET Core 3.1 on Visual Studio 2019
* Azure functions v3, Microsoft Graph API and Send Grid API

## Code snippets
### Retrieve the data in excel sheet using MS Graph API
```
public List<Expense> RetrieveData(string file, string worksheet, string table, ILogger log)
{
    string rowcontent = string.Empty;
    List<Expense> expenses = new List<Expense>();
    List<Expense> filteredExpenses = new List<Expense>();

    var client = new RestClient(string.Format("{0}me/drive/root/children/{1}/workbook/worksheets/{2}/Tables/{3}/rows", 
      baseurl, file, worksheet, table));
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
            Expense expense = new Expense { Description = row[1], Price = Convert.ToDouble(row[2]),
              DueDate = string.IsNullOrEmpty(row[3]) ? null : (DateTime?)DateTime.FromOADate(Convert.ToDouble(row[3]))};
            expenses.Add(expense);
        }

        DateTime fromDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
        DateTime toDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 
          DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month));

        filteredExpenses = expenses.Where(p => p.DueDate >= fromDate && p.DueDate < toDate).ToList();

    }
    return filteredExpenses;
}
```

### Send an email using Send Grid API
```
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
```

### Build the email message
```
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
```

### Configure the Sendgrid client
```
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
```

### Settings in the application
```
"AccessToken": "",
"emailServer": "smtp.sendgrid.net",
"emailPort": "587",
"emailCredentialUserName": "",
 "emailCredentialPassword": "",
 "fromAddress": ""
```
