namespace AFH.Sharepoint.Integration.Tests.Helpers;

[AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
public sealed class OppAttribute : Attribute
{
    public string Name { get; }
    public OppAttribute(string name) => Name = name;
}
