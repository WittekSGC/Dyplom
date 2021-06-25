namespace Dyplom.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class Managements : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.LeadTeachers",
                c => new
                    {
                        id = c.Int(nullable: false, identity: true),
                        classid = c.Int(nullable: false),
                        Fullname = c.String(),
                        teacherlogin = c.String(),
                        teacherpassword = c.String(),
                    })
                .PrimaryKey(t => t.id);
            
            CreateTable(
                "dbo.Management",
                c => new
                    {
                        managementid = c.Int(nullable: false, identity: true),
                        ManagementLogin = c.String(),
                        ManagementPassword = c.String(),
                    })
                .PrimaryKey(t => t.managementid);
            
            CreateTable(
                "dbo.Students",
                c => new
                    {
                        id = c.Int(nullable: false, identity: true),
                        studentName = c.String(),
                        homeAdressReg = c.String(),
                        homeAdressRel = c.String(),
                        studentTel = c.String(),
                        motherName = c.String(),
                        motherPlaceOfWork = c.String(),
                        motherWorkPhone = c.String(),
                        motherMobPhone = c.String(),
                        fatherName = c.String(),
                        fatherPlaceOfWork = c.String(),
                        fatherWorkPhone = c.String(),
                        fatherMobPhone = c.String(),
                        isChildInvalit = c.Boolean(nullable: false),
                        isChildWithOPFR = c.Boolean(nullable: false),
                        childInCustody = c.Boolean(nullable: false),
                        isChildInFosterCare = c.Boolean(nullable: false),
                        doesChildStudyAtHome = c.Boolean(nullable: false),
                        isChildRegistered = c.Boolean(nullable: false),
                        numberOfChildInFamilyUnder18 = c.Int(nullable: false),
                        incompleteFamilyOneMother = c.Boolean(nullable: false),
                        incompleteFamilyOneFather = c.Boolean(nullable: false),
                        aSingleMother = c.Boolean(nullable: false),
                        motherEducation = c.String(),
                        fatherEducation = c.String(),
                        motherStatus = c.String(),
                        fatherStatus = c.String(),
                        classid = c.Int(nullable: false),
                    })
                .PrimaryKey(t => t.id);
            
            CreateTable(
                "dbo.Users",
                c => new
                    {
                        UserId = c.Int(nullable: false, identity: true),
                        UserLogin = c.String(),
                        UserPassword = c.String(),
                    })
                .PrimaryKey(t => t.UserId);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.Users");
            DropTable("dbo.Students");
            DropTable("dbo.Management");
            DropTable("dbo.LeadTeachers");
        }
    }
}
