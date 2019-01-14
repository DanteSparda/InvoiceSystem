using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InvoiceSystem.Models
{
  public class InvoiceLineModel
  {
    public string Code { get; set; }

    public string Description { get; set; }

    public string UoM { get; set; }

    public string Quantity { get; set; }

    public decimal TotalPrice { get; set; }

    public decimal UnitPrice { get; set; }

  }
}
