using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InvoiceSystem.Models
{
  public class Invoice
  {
    public Invoice()
    {
      Lines = new List<InvoiceLineModel>();
    }

    public InvoiceHeaderModel Header { get; set; }

    public List<InvoiceLineModel> Lines { get; set; }

    public string Name { get; set; }

  }
}
