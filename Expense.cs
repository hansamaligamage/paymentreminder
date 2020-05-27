using System;
using System.Collections.Generic;
using System.Text;

namespace paymentreminder
{
    public class Expense
    {
        public string Description { get; set; }
        public double Price { get; set; }
        public DateTime? DueDate { get; set; }
    }
}
