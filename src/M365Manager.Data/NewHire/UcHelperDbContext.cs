using Microsoft.EntityFrameworkCore;

namespace M365Manager.Data.NewHire;

/// <summary>
/// Read/write access to the legacy telephony database (historically "uchelper"), which lives on the
/// same SQL Server as the app's own database but under a different catalog - hence its own context
/// and its own connection string provider (<see cref="INewHireConnectionStringProvider"/>).
///
/// The schema is not ours: table and column names are lower-case, mixed-case and inconsistent
/// because they were grown by hand and then scaffolded. Every mapping below is explicit for exactly
/// that reason - the CLR names follow this project's conventions and the column names follow the
/// database. Nothing here creates or migrates anything; the tables are owned by the telephony
/// imports and this context only ever attaches to what is already there.
/// </summary>
public sealed class UcHelperDbContext : DbContext
{
    public UcHelperDbContext(DbContextOptions<UcHelperDbContext> options)
        : base(options)
    {
    }

    public DbSet<UcLocation> Locations => Set<UcLocation>();
    public DbSet<UcDidRange> DidRanges => Set<UcDidRange>();
    public DbSet<UcEndpoint> Endpoints => Set<UcEndpoint>();
    public DbSet<UcBlockedDid> BlockedDids => Set<UcBlockedDid>();
    public DbSet<UcE3License> E3Licenses => Set<UcE3License>();
    public DbSet<UcNewHireLog> NewHireLogs => Set<UcNewHireLog>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        var location = modelBuilder.Entity<UcLocation>();
        location.ToTable("locationconfiguration");
        location.HasKey(x => x.Id);
        location.Property(x => x.Id).HasColumnName("ID");
        location.Property(x => x.Name).HasColumnName("Name");
        location.Property(x => x.LocationCode).HasColumnName("LocationCode");
        location.Property(x => x.City).HasColumnName("City");
        location.Property(x => x.Country).HasColumnName("Country");
        location.Property(x => x.TeamsVoiceEnabled).HasColumnName("TeamsVoiceEnabled");

        var did = modelBuilder.Entity<UcDidRange>();
        did.ToTable("did");
        did.HasKey(x => x.Id);
        did.Property(x => x.Id).HasColumnName("ID");
        did.Property(x => x.DidStart).HasColumnName("DIDSTART");
        did.Property(x => x.DidEnd).HasColumnName("DIDEND");
        did.Property(x => x.LocationCode).HasColumnName("LocationCode");
        did.Property(x => x.Utilization).HasColumnName("utilization");
        did.Property(x => x.Notes).HasColumnName("Notes");
        did.Property(x => x.Sdap).HasColumnName("SDAP");
        did.Property(x => x.Provider).HasColumnName("PROVIDER");

        var endpoint = modelBuilder.Entity<UcEndpoint>();
        endpoint.ToTable("endpoints");
        endpoint.HasKey(x => x.Id);
        endpoint.Property(x => x.Id).HasColumnName("ID");
        endpoint.Property(x => x.DisplayName).HasColumnName("DisplayName");
        endpoint.Property(x => x.SamAccountName).HasColumnName("SamAccountName");
        endpoint.Property(x => x.SipAddress).HasColumnName("SipAddress");
        endpoint.Property(x => x.LineUri).HasColumnName("LineURI");
        endpoint.Property(x => x.Did).HasColumnName("DID");
        endpoint.Property(x => x.Type).HasColumnName("Type");
        endpoint.Property(x => x.VoicePolicy).HasColumnName("VoicePolicy");
        endpoint.Property(x => x.DialPlan).HasColumnName("DialPlan");
        endpoint.Property(x => x.EnterpriseVoiceEnabled).HasColumnName("EnterpriseVoiceEnabled");
        endpoint.Property(x => x.HostedVoiceMail).HasColumnName("HostedVoiceMail");
        endpoint.Property(x => x.EmergencyCallingPolicy).HasColumnName("EmergencyCallingPolicy");
        endpoint.Property(x => x.EmergencyCallRoutingPolicy).HasColumnName("EmergencyCallRoutingPolicy");
        endpoint.Property(x => x.Office).HasColumnName("Office");

        var blocked = modelBuilder.Entity<UcBlockedDid>();
        blocked.ToTable("blockeddids");
        blocked.HasKey(x => x.Id);
        blocked.Property(x => x.Id).HasColumnName("ID");
        blocked.Property(x => x.Did).HasColumnName("did");
        blocked.Property(x => x.Notes).HasColumnName("notes");

        var license = modelBuilder.Entity<UcE3License>();
        license.ToTable("e3licensed");
        license.HasKey(x => x.Id);
        license.Property(x => x.Id).HasColumnName("ID");
        license.Property(x => x.SamAccountName).HasColumnName("samaccountname");
        license.Property(x => x.E3Licensed).HasColumnName("e3licensed");
        license.Property(x => x.HasPhoneLicense).HasColumnName("HasPhoneLicense");

        var log = modelBuilder.Entity<UcNewHireLog>();
        log.ToTable("newhirelog");
        log.HasKey(x => x.NewHireLogId);
        log.Property(x => x.NewHireLogId).HasColumnName("newhirelogid");
        log.Property(x => x.SamAccountName).HasColumnName("samaccountname");
        log.Property(x => x.AdminAccount).HasColumnName("adminaccount");
        log.Property(x => x.Date).HasColumnName("date").HasColumnType("datetime");
        log.Property(x => x.SnTicket).HasColumnName("snticket");
        log.Property(x => x.AllVariablesSet).HasColumnName("AllVariablesSet");
        log.Property(x => x.CrdFileExisting).HasColumnName("CrdFileExisting");
        log.Property(x => x.CrdFileUpToDate).HasColumnName("CrdFileUpToDate");
        log.Property(x => x.TeamsConnected).HasColumnName("TeamsConnected");
        log.Property(x => x.TeamsConnectedError).HasColumnName("TeamsConnectedError");
        log.Property(x => x.MsolConnected).HasColumnName("MsolConnected");
        log.Property(x => x.MsolConnectedError).HasColumnName("MsolConnectedError");
        log.Property(x => x.MsRtcSipDeploymentLocator).HasColumnName("msRTCSIPDeploymentLocator");
        log.Property(x => x.MsRtcSipDeploymentLocatorError).HasColumnName("msRTCSIPDeploymentLocatorError");
        log.Property(x => x.MsRtcSipFederationEnabled).HasColumnName("msRTCSIPFederationEnabled");
        log.Property(x => x.MsRtcSipFederationEnabledError).HasColumnName("msRTCSIPFederationEnabledError");
        log.Property(x => x.MsRtcSipInternetAccessEnabled).HasColumnName("msRTCSIPInternetAccessEnabled");
        log.Property(x => x.MsRtcSipInternetAccessEnabledError).HasColumnName("msRTCSIPInternetAccessEnabledError");
        log.Property(x => x.MsRtcSipUserEnabled).HasColumnName("msRTCSIPUserEnabled");
        log.Property(x => x.MsRtcSipUserEnabledError).HasColumnName("msRTCSIPUserEnabledError");
        log.Property(x => x.MsRtcSipLine).HasColumnName("msRTCSIPLine");
        log.Property(x => x.MsRtcSipLineError).HasColumnName("msRTCSIPLineError");
        log.Property(x => x.MsRtcSipPrimaryUserAddress).HasColumnName("msRTCSIPPrimaryUserAddress");
        log.Property(x => x.MsRtcSipPrimaryUserAddressError).HasColumnName("msRTCSIPPrimaryUserAddressError");
        log.Property(x => x.VoiceRoutingPolicy).HasColumnName("VoiceRoutingPolicy");
        log.Property(x => x.VoiceRoutingPolicyError).HasColumnName("VoiceRoutingPolicyError");
        log.Property(x => x.DialPlan).HasColumnName("DialPlan");
        log.Property(x => x.DialPlanError).HasColumnName("DialPlanError");
        log.Property(x => x.EmergencyCallingPolicy).HasColumnName("EmergencyCallingPolicy");
        log.Property(x => x.EmergencyCallingPolicyError).HasColumnName("EmergencyCallingPolicyError");
        log.Property(x => x.EmergencyCallRoutingPolicy).HasColumnName("EmergencyCallRoutingPolicy");
        log.Property(x => x.EmergencyCallRoutingPolicyError).HasColumnName("EmergencyCallRoutingPolicyError");
        log.Property(x => x.TeamsUpgradePolicy).HasColumnName("TeamsUpgradePolicy");
        log.Property(x => x.TeamsUpgradePolicyError).HasColumnName("TeamsUpgradePolicyError");
        log.Property(x => x.EvEnabled).HasColumnName("EVEnabled");
        log.Property(x => x.EvEnabledError).HasColumnName("EVEnabledError");
        log.Property(x => x.SqlServerModule).HasColumnName("sqlservermodule");
        log.Property(x => x.PhoneLicense).HasColumnName("PhoneLicense");
        log.Property(x => x.PhoneLicenseError).HasColumnName("PhoneLicenseError");
    }
}
