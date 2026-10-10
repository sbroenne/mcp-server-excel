namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>MSOLAP interactive authentication setting, not a command to sign in.</summary>
public enum ConnectionInteractiveLogin
{
    /// <summary>Let the provider and other connection settings determine sign-in behavior.</summary>
    Default,
    /// <summary>Allow interactive sign-in when silent authentication fails.</summary>
    Enabled,
    /// <summary>Do not allow interactive sign-in as a fallback.</summary>
    Disabled,
    /// <summary>Request interactive sign-in instead of silent authentication.</summary>
    Always
}

/// <summary>MSOLAP identity selection when no explicit User ID overrides it.</summary>
public enum ConnectionIdentityMode
{
    /// <summary>Let the provider and system configuration determine identity selection.</summary>
    Default,
    /// <summary>Use the current Windows user's identity.</summary>
    CurrentUser,
    /// <summary>Reuse the selected data-source identity for the application process lifetime.</summary>
    Connection,
    /// <summary>Reuse the identity selected for the first connection in the process.</summary>
    Process
}
