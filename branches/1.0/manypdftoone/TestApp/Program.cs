using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestApp
{
    class Program
    {
        static void Main(string[] args)
        {
            manypdftoone.ManyPdfToOne mpto = new manypdftoone.ManyPdfToOne(@"C:\Users\pbrzezinski\Downloads", @"C:\Users\pbrzezinski\Documents\MPdf", manypdftoone.eMergePagesMode.TwoSide);
            mpto.Merge();
        }
    }
}
