namespace AFH.Common.SharePointUtils.Connector;

public class ConnectorResult<T>
{
    public bool Success { get; }
    public T? Data { get; }
    public Exception? Error { get; }

    private ConnectorResult(bool success, T? data, Exception? error)
    {
        Success = success;
        Data = data;
        Error = error;
    }

    public static ConnectorResult<T> Ok(T data) => new(true, data, null);
    public static ConnectorResult<T> Fail(Exception ex) => new(false, default, ex);
}
