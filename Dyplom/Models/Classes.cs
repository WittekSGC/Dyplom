using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dyplom.Models
{
    class Classes
    {
        [Key]
        public int classid { get; set; }
        public int StudyYear { get; set; }
        public string GradeSymbol { get; set; }
    }
}
