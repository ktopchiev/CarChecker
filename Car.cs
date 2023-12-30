using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CarChecker
{
    public class Car
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public string AssemblyYear { get; set; }
        public int Kilometers { get; set; }
        public string FuelType { get; set; }
        public double Price { get; set; }
        public string Url { get; set; }
    }
}
