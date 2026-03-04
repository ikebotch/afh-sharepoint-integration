namespace AFH.Sharepoint.Integration.Tests.Helpers;


[AttributeUsage(AttributeTargets.Class | AttributeTargets.Method, AllowMultiple = false)]
public sealed class KindAttribute : Attribute
{
    public string Name { get; }
    public KindAttribute(string name) => Name = name;
}
