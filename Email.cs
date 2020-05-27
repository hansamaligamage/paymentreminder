using System;
using System.Collections.Generic;
using System.Text;

namespace paymentreminder
{
    public class Email
    {
        public string Subject { get; set; }
        public string MailRecipientsTo { get; set; }
        public string Content { get; set; }
    }
}
