namespace KPSZI.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class first : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.ConfigOptions",
                c => new
                    {
                        ConfigOptionId = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                        Description = c.String(),
                        DefenceClass = c.String(),
                        GISMeasure_GISMeasureId = c.Int(),
                    })
                .PrimaryKey(t => t.ConfigOptionId)
                .ForeignKey("dbo.GISMeasures", t => t.GISMeasure_GISMeasureId)
                .Index(t => t.GISMeasure_GISMeasureId);
            
            CreateTable(
                "dbo.GISMeasures",
                c => new
                    {
                        GISMeasureId = c.Int(nullable: false, identity: true),
                        Number = c.Int(nullable: false),
                        Description = c.String(),
                        MinimalRequirementDefenceClass = c.Int(nullable: false),
                        isOnlyISPDn = c.Boolean(nullable: false),
                        MeasureGroup_MeasureGroupId = c.Int(),
                        SZI_SZIId = c.Int(),
                    })
                .PrimaryKey(t => t.GISMeasureId)
                .ForeignKey("dbo.MeasureGroups", t => t.MeasureGroup_MeasureGroupId)
                .ForeignKey("dbo.SZIs", t => t.SZI_SZIId)
                .Index(t => t.MeasureGroup_MeasureGroupId)
                .Index(t => t.SZI_SZIId);
            
            CreateTable(
                "dbo.MeasureGroups",
                c => new
                    {
                        MeasureGroupId = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                        ShortName = c.String(),
                        Description = c.String(),
                    })
                .PrimaryKey(t => t.MeasureGroupId);
            
            CreateTable(
                "dbo.ISPDNMeasures",
                c => new
                    {
                        ISPDNMeasureId = c.Int(nullable: false, identity: true),
                        Number = c.Int(nullable: false),
                        Description = c.String(),
                        MinimalRequirementDefenceLevel = c.Int(nullable: false),
                        MeasureGroup_MeasureGroupId = c.Int(),
                    })
                .PrimaryKey(t => t.ISPDNMeasureId)
                .ForeignKey("dbo.MeasureGroups", t => t.MeasureGroup_MeasureGroupId)
                .Index(t => t.MeasureGroup_MeasureGroupId);
            
            CreateTable(
                "dbo.SZIs",
                c => new
                    {
                        SZIId = c.Int(nullable: false, identity: true),
                        DefenceClass = c.String(),
                        Name = c.String(),
                        Certificate = c.String(),
                        DateOfEnd = c.DateTime(nullable: false),
                        NDVControlLevel = c.Int(nullable: false),
                        SVTClass = c.String(),
                        TU = c.String(),
                        CPUFrequencyRequirements = c.Int(nullable: false),
                        CPUCoresRequirements = c.Int(nullable: false),
                        MemoryRequirements = c.Int(nullable: false),
                        DiscSpaceRequirements = c.Int(nullable: false),
                        ISPDNMeasure_ISPDNMeasureId = c.Int(),
                    })
                .PrimaryKey(t => t.SZIId)
                .ForeignKey("dbo.ISPDNMeasures", t => t.ISPDNMeasure_ISPDNMeasureId)
                .Index(t => t.ISPDNMeasure_ISPDNMeasureId);
            
            CreateTable(
                "dbo.SZISorts",
                c => new
                    {
                        SZISortId = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                        ShortName = c.String(),
                        Number = c.Int(nullable: false),
                    })
                .PrimaryKey(t => t.SZISortId);
            
            CreateTable(
                "dbo.SFHs",
                c => new
                    {
                        SFHId = c.Int(nullable: false, identity: true),
                        SFHNumber = c.Int(nullable: false),
                        Name = c.String(),
                        ProjectSecurity = c.Int(nullable: false),
                        SFHType_SFHTypeId = c.Int(),
                    })
                .PrimaryKey(t => t.SFHId)
                .ForeignKey("dbo.SFHTypes", t => t.SFHType_SFHTypeId)
                .Index(t => t.SFHType_SFHTypeId);
            
            CreateTable(
                "dbo.SFHTypes",
                c => new
                    {
                        SFHTypeId = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                        MultipleChoice = c.Boolean(nullable: false),
                    })
                .PrimaryKey(t => t.SFHTypeId);
            
            CreateTable(
                "dbo.Threats",
                c => new
                    {
                        ThreatId = c.Int(nullable: false, identity: true),
                        ThreatNumber = c.Int(nullable: false),
                        Name = c.String(),
                        Description = c.String(),
                        ObjectOfInfluence = c.String(),
                        ConfidenceViolation = c.Boolean(nullable: false),
                        IntegrityViolation = c.Boolean(nullable: false),
                        AvailabilityViolation = c.Boolean(nullable: false),
                        DateOfAdd = c.DateTime(nullable: false),
                        DateOfChange = c.DateTime(nullable: false),
                    })
                .PrimaryKey(t => t.ThreatId);
            
            CreateTable(
                "dbo.ImplementWays",
                c => new
                    {
                        ImplementWayId = c.Int(nullable: false, identity: true),
                        WayNumber = c.Int(nullable: false),
                        WayName = c.String(),
                    })
                .PrimaryKey(t => t.ImplementWayId);
            
            CreateTable(
                "dbo.ThreatSources",
                c => new
                    {
                        ThreatSourceId = c.Int(nullable: false, identity: true),
                        InternalIntruder = c.Boolean(nullable: false),
                        Potencial = c.Int(nullable: false),
                    })
                .PrimaryKey(t => t.ThreatSourceId);
            
            CreateTable(
                "dbo.Vulnerabilities",
                c => new
                    {
                        VulnerabilityId = c.Int(nullable: false, identity: true),
                        VulnerabilityNumber = c.Int(nullable: false),
                        VulnerabilityName = c.String(),
                        VulnerabilityDescription = c.String(),
                    })
                .PrimaryKey(t => t.VulnerabilityId);
            
            CreateTable(
                "dbo.InfoTypes",
                c => new
                    {
                        InfoTypeId = c.Int(nullable: false, identity: true),
                        TypeName = c.String(),
                    })
                .PrimaryKey(t => t.InfoTypeId);
            
            CreateTable(
                "dbo.IntruderTypes",
                c => new
                    {
                        IntruderTypeId = c.Int(nullable: false, identity: true),
                        TypeName = c.String(),
                    })
                .PrimaryKey(t => t.IntruderTypeId);
            
            CreateTable(
                "dbo.TCUIs",
                c => new
                    {
                        TCUIId = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                        TCUIType_TCUITypeId = c.Int(),
                    })
                .PrimaryKey(t => t.TCUIId)
                .ForeignKey("dbo.TCUITypes", t => t.TCUIType_TCUITypeId)
                .Index(t => t.TCUIType_TCUITypeId);
            
            CreateTable(
                "dbo.TCUIThreats",
                c => new
                    {
                        TCUIThreatId = c.Int(nullable: false, identity: true),
                        Identificator = c.String(),
                        Name = c.String(),
                        Description = c.String(),
                        TCUI_TCUIId = c.Int(),
                    })
                .PrimaryKey(t => t.TCUIThreatId)
                .ForeignKey("dbo.TCUIs", t => t.TCUI_TCUIId)
                .Index(t => t.TCUI_TCUIId);
            
            CreateTable(
                "dbo.TCUITypes",
                c => new
                    {
                        TCUITypeId = c.Int(nullable: false, identity: true),
                        Name = c.String(),
                    })
                .PrimaryKey(t => t.TCUITypeId);
            
            CreateTable(
                "dbo.TechnogenicMeasures",
                c => new
                    {
                        TechnogenicMeasureId = c.Int(nullable: false, identity: true),
                        Description = c.String(),
                    })
                .PrimaryKey(t => t.TechnogenicMeasureId);
            
            CreateTable(
                "dbo.TechnogenicThreats",
                c => new
                    {
                        TechnogenicThreatId = c.Int(nullable: false, identity: true),
                        Identificator = c.String(),
                        Description = c.String(),
                    })
                .PrimaryKey(t => t.TechnogenicThreatId);
            
            CreateTable(
                "dbo.SZISortGISMeasures",
                c => new
                    {
                        SZISort_SZISortId = c.Int(nullable: false),
                        GISMeasure_GISMeasureId = c.Int(nullable: false),
                    })
                .PrimaryKey(t => new { t.SZISort_SZISortId, t.GISMeasure_GISMeasureId })
                .ForeignKey("dbo.SZISorts", t => t.SZISort_SZISortId, cascadeDelete: true)
                .ForeignKey("dbo.GISMeasures", t => t.GISMeasure_GISMeasureId, cascadeDelete: true)
                .Index(t => t.SZISort_SZISortId)
                .Index(t => t.GISMeasure_GISMeasureId);
            
            CreateTable(
                "dbo.SZISortSZIs",
                c => new
                    {
                        SZISort_SZISortId = c.Int(nullable: false),
                        SZI_SZIId = c.Int(nullable: false),
                    })
                .PrimaryKey(t => new { t.SZISort_SZISortId, t.SZI_SZIId })
                .ForeignKey("dbo.SZISorts", t => t.SZISort_SZISortId, cascadeDelete: true)
                .ForeignKey("dbo.SZIs", t => t.SZI_SZIId, cascadeDelete: true)
                .Index(t => t.SZISort_SZISortId)
                .Index(t => t.SZI_SZIId);
            
            CreateTable(
                "dbo.SFHGISMeasures",
                c => new
                    {
                        SFH_SFHId = c.Int(nullable: false),
                        GISMeasure_GISMeasureId = c.Int(nullable: false),
                    })
                .PrimaryKey(t => new { t.SFH_SFHId, t.GISMeasure_GISMeasureId })
                .ForeignKey("dbo.SFHs", t => t.SFH_SFHId, cascadeDelete: true)
                .ForeignKey("dbo.GISMeasures", t => t.GISMeasure_GISMeasureId, cascadeDelete: true)
                .Index(t => t.SFH_SFHId)
                .Index(t => t.GISMeasure_GISMeasureId);
            
            CreateTable(
                "dbo.ThreatGISMeasures",
                c => new
                    {
                        Threat_ThreatId = c.Int(nullable: false),
                        GISMeasure_GISMeasureId = c.Int(nullable: false),
                    })
                .PrimaryKey(t => new { t.Threat_ThreatId, t.GISMeasure_GISMeasureId })
                .ForeignKey("dbo.Threats", t => t.Threat_ThreatId, cascadeDelete: true)
                .ForeignKey("dbo.GISMeasures", t => t.GISMeasure_GISMeasureId, cascadeDelete: true)
                .Index(t => t.Threat_ThreatId)
                .Index(t => t.GISMeasure_GISMeasureId);
            
            CreateTable(
                "dbo.ImplementWayThreats",
                c => new
                    {
                        ImplementWay_ImplementWayId = c.Int(nullable: false),
                        Threat_ThreatId = c.Int(nullable: false),
                    })
                .PrimaryKey(t => new { t.ImplementWay_ImplementWayId, t.Threat_ThreatId })
                .ForeignKey("dbo.ImplementWays", t => t.ImplementWay_ImplementWayId, cascadeDelete: true)
                .ForeignKey("dbo.Threats", t => t.Threat_ThreatId, cascadeDelete: true)
                .Index(t => t.ImplementWay_ImplementWayId)
                .Index(t => t.Threat_ThreatId);
            
            CreateTable(
                "dbo.ThreatSFHs",
                c => new
                    {
                        Threat_ThreatId = c.Int(nullable: false),
                        SFH_SFHId = c.Int(nullable: false),
                    })
                .PrimaryKey(t => new { t.Threat_ThreatId, t.SFH_SFHId })
                .ForeignKey("dbo.Threats", t => t.Threat_ThreatId, cascadeDelete: true)
                .ForeignKey("dbo.SFHs", t => t.SFH_SFHId, cascadeDelete: true)
                .Index(t => t.Threat_ThreatId)
                .Index(t => t.SFH_SFHId);
            
            CreateTable(
                "dbo.ThreatSourceThreats",
                c => new
                    {
                        ThreatSource_ThreatSourceId = c.Int(nullable: false),
                        Threat_ThreatId = c.Int(nullable: false),
                    })
                .PrimaryKey(t => new { t.ThreatSource_ThreatSourceId, t.Threat_ThreatId })
                .ForeignKey("dbo.ThreatSources", t => t.ThreatSource_ThreatSourceId, cascadeDelete: true)
                .ForeignKey("dbo.Threats", t => t.Threat_ThreatId, cascadeDelete: true)
                .Index(t => t.ThreatSource_ThreatSourceId)
                .Index(t => t.Threat_ThreatId);
            
            CreateTable(
                "dbo.VulnerabilityThreats",
                c => new
                    {
                        Vulnerability_VulnerabilityId = c.Int(nullable: false),
                        Threat_ThreatId = c.Int(nullable: false),
                    })
                .PrimaryKey(t => new { t.Vulnerability_VulnerabilityId, t.Threat_ThreatId })
                .ForeignKey("dbo.Vulnerabilities", t => t.Vulnerability_VulnerabilityId, cascadeDelete: true)
                .ForeignKey("dbo.Threats", t => t.Threat_ThreatId, cascadeDelete: true)
                .Index(t => t.Vulnerability_VulnerabilityId)
                .Index(t => t.Threat_ThreatId);
            
        }
        
        public override void Down()
        {
            DropForeignKey("dbo.TCUIs", "TCUIType_TCUITypeId", "dbo.TCUITypes");
            DropForeignKey("dbo.TCUIThreats", "TCUI_TCUIId", "dbo.TCUIs");
            DropForeignKey("dbo.VulnerabilityThreats", "Threat_ThreatId", "dbo.Threats");
            DropForeignKey("dbo.VulnerabilityThreats", "Vulnerability_VulnerabilityId", "dbo.Vulnerabilities");
            DropForeignKey("dbo.ThreatSourceThreats", "Threat_ThreatId", "dbo.Threats");
            DropForeignKey("dbo.ThreatSourceThreats", "ThreatSource_ThreatSourceId", "dbo.ThreatSources");
            DropForeignKey("dbo.ThreatSFHs", "SFH_SFHId", "dbo.SFHs");
            DropForeignKey("dbo.ThreatSFHs", "Threat_ThreatId", "dbo.Threats");
            DropForeignKey("dbo.ImplementWayThreats", "Threat_ThreatId", "dbo.Threats");
            DropForeignKey("dbo.ImplementWayThreats", "ImplementWay_ImplementWayId", "dbo.ImplementWays");
            DropForeignKey("dbo.ThreatGISMeasures", "GISMeasure_GISMeasureId", "dbo.GISMeasures");
            DropForeignKey("dbo.ThreatGISMeasures", "Threat_ThreatId", "dbo.Threats");
            DropForeignKey("dbo.SFHs", "SFHType_SFHTypeId", "dbo.SFHTypes");
            DropForeignKey("dbo.SFHGISMeasures", "GISMeasure_GISMeasureId", "dbo.GISMeasures");
            DropForeignKey("dbo.SFHGISMeasures", "SFH_SFHId", "dbo.SFHs");
            DropForeignKey("dbo.SZIs", "ISPDNMeasure_ISPDNMeasureId", "dbo.ISPDNMeasures");
            DropForeignKey("dbo.SZISortSZIs", "SZI_SZIId", "dbo.SZIs");
            DropForeignKey("dbo.SZISortSZIs", "SZISort_SZISortId", "dbo.SZISorts");
            DropForeignKey("dbo.SZISortGISMeasures", "GISMeasure_GISMeasureId", "dbo.GISMeasures");
            DropForeignKey("dbo.SZISortGISMeasures", "SZISort_SZISortId", "dbo.SZISorts");
            DropForeignKey("dbo.GISMeasures", "SZI_SZIId", "dbo.SZIs");
            DropForeignKey("dbo.ISPDNMeasures", "MeasureGroup_MeasureGroupId", "dbo.MeasureGroups");
            DropForeignKey("dbo.GISMeasures", "MeasureGroup_MeasureGroupId", "dbo.MeasureGroups");
            DropForeignKey("dbo.ConfigOptions", "GISMeasure_GISMeasureId", "dbo.GISMeasures");
            DropIndex("dbo.VulnerabilityThreats", new[] { "Threat_ThreatId" });
            DropIndex("dbo.VulnerabilityThreats", new[] { "Vulnerability_VulnerabilityId" });
            DropIndex("dbo.ThreatSourceThreats", new[] { "Threat_ThreatId" });
            DropIndex("dbo.ThreatSourceThreats", new[] { "ThreatSource_ThreatSourceId" });
            DropIndex("dbo.ThreatSFHs", new[] { "SFH_SFHId" });
            DropIndex("dbo.ThreatSFHs", new[] { "Threat_ThreatId" });
            DropIndex("dbo.ImplementWayThreats", new[] { "Threat_ThreatId" });
            DropIndex("dbo.ImplementWayThreats", new[] { "ImplementWay_ImplementWayId" });
            DropIndex("dbo.ThreatGISMeasures", new[] { "GISMeasure_GISMeasureId" });
            DropIndex("dbo.ThreatGISMeasures", new[] { "Threat_ThreatId" });
            DropIndex("dbo.SFHGISMeasures", new[] { "GISMeasure_GISMeasureId" });
            DropIndex("dbo.SFHGISMeasures", new[] { "SFH_SFHId" });
            DropIndex("dbo.SZISortSZIs", new[] { "SZI_SZIId" });
            DropIndex("dbo.SZISortSZIs", new[] { "SZISort_SZISortId" });
            DropIndex("dbo.SZISortGISMeasures", new[] { "GISMeasure_GISMeasureId" });
            DropIndex("dbo.SZISortGISMeasures", new[] { "SZISort_SZISortId" });
            DropIndex("dbo.TCUIThreats", new[] { "TCUI_TCUIId" });
            DropIndex("dbo.TCUIs", new[] { "TCUIType_TCUITypeId" });
            DropIndex("dbo.SFHs", new[] { "SFHType_SFHTypeId" });
            DropIndex("dbo.SZIs", new[] { "ISPDNMeasure_ISPDNMeasureId" });
            DropIndex("dbo.ISPDNMeasures", new[] { "MeasureGroup_MeasureGroupId" });
            DropIndex("dbo.GISMeasures", new[] { "SZI_SZIId" });
            DropIndex("dbo.GISMeasures", new[] { "MeasureGroup_MeasureGroupId" });
            DropIndex("dbo.ConfigOptions", new[] { "GISMeasure_GISMeasureId" });
            DropTable("dbo.VulnerabilityThreats");
            DropTable("dbo.ThreatSourceThreats");
            DropTable("dbo.ThreatSFHs");
            DropTable("dbo.ImplementWayThreats");
            DropTable("dbo.ThreatGISMeasures");
            DropTable("dbo.SFHGISMeasures");
            DropTable("dbo.SZISortSZIs");
            DropTable("dbo.SZISortGISMeasures");
            DropTable("dbo.TechnogenicThreats");
            DropTable("dbo.TechnogenicMeasures");
            DropTable("dbo.TCUITypes");
            DropTable("dbo.TCUIThreats");
            DropTable("dbo.TCUIs");
            DropTable("dbo.IntruderTypes");
            DropTable("dbo.InfoTypes");
            DropTable("dbo.Vulnerabilities");
            DropTable("dbo.ThreatSources");
            DropTable("dbo.ImplementWays");
            DropTable("dbo.Threats");
            DropTable("dbo.SFHTypes");
            DropTable("dbo.SFHs");
            DropTable("dbo.SZISorts");
            DropTable("dbo.SZIs");
            DropTable("dbo.ISPDNMeasures");
            DropTable("dbo.MeasureGroups");
            DropTable("dbo.GISMeasures");
            DropTable("dbo.ConfigOptions");
        }
    }
}
