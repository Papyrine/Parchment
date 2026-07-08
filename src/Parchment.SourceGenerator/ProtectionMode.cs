/// <summary>
/// SG-side copy of <c>Parchment.ProtectionMode</c> — the generator cannot reference
/// Parchment.dll. Values must stay in lockstep; the attribute's named argument arrives as the
/// underlying int and emission writes the runtime enum by member name.
/// </summary>
enum ProtectionMode
{
    WhenEditable,
    None
}
