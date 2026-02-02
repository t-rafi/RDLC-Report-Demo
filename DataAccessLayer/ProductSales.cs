using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccessLayer.Models
{
    public class ProductSales
    {
        public string ProdCat { get; set; }
        public string SubCat { get; set; }
        public int OrderYear { get; set; }
        public string OrderQtr { get; set; }
        public decimal Sales { get; set; }

        //new 2
        public string Region { get; set; }
        public string Manager { get; set; }
    }
}
