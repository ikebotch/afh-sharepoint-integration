namespace AFH.Sharepoint.Integration.Tests.Helpers;

[AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
public sealed class AreaAttribute : Attribute
{
    public string Name { get; }
    public AreaAttribute(string name) => Name = name;
}
