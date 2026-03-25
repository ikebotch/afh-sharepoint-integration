namespace AFH.Common.SharePointUtils.Exceptions;

public sealed class SharePointSdkException : Exception
{
    public SharePointSdkException(string message, Exception? innerException = null)
        : base(message, innerException)
    {
    }
}
