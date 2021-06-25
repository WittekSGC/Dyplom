using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dyplom.Models
{
    class LeadTeachers
    {
        public int id { get; set; }
        public int classid { get; set; }
        public string Fullname { get; set; }
        public string teacherlogin { get; set; }
        public string teacherpassword { get; set; }
    }
}
