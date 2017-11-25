using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DNRegisterListFilter
{
    public class Profile
    {
        public Profile()
        {
            this.Count = 1;
        }

        public string Name { get; set; }
        public string Email { get; set; }
        public string School { get; set; }
        public string Major { get; set; }
        public string Grade { get; set; }
        public int Count { get; set; }
    }
}
