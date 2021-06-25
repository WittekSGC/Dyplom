using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dyplom.Models
{
    class Students
    {
        public int id { get; set; }
        public string studentName { get; set; }
        public DateTime birthdate { get; set; }
        public string homeAdressReg { get; set; }
        public string homeAdressRel { get; set; }
        public string studentTel { get; set; }
        public string motherName { get; set; }
        public string motherPlaceOfWork { get; set; }
        public string motherWorkPhone { get; set; }
        public string motherMobPhone { get; set; }
        public string fatherName { get; set; }
        public string fatherPlaceOfWork { get; set; }
        public string fatherWorkPhone { get; set; }
        public string fatherMobPhone { get; set; }
        public bool isChildInvalit { get; set; }
        public bool isChildWithOPFR { get; set; }
        public bool childInCustody { get; set; }
        public bool isChildInFosterCare { get; set; }
        public bool doesChildStudyAtHome { get; set; }
        public bool isChildRegistered { get; set; }
        public int numberOfChildInFamilyUnder18 { get; set; }
        public bool incompleteFamilyOneMother { get; set; }
        public bool incompleteFamilyOneFather { get; set; }
        public bool aSingleMother { get; set; }
        public string motherEducation { get; set; }
        public string fatherEducation { get; set; }
        public string motherStatus { get; set; }
        public string fatherStatus { get; set; }
        public int classid { get; set; }

    }
}
