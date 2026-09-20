namespace AudioDeviceSwitcher;

public enum AppRouteState { Matched, WaitingForSession, Unknown, Drifted }

public static class AppRouteStatus
{
    public static AppRouteState Evaluate(bool hasSession, bool querySucceeded,
        string? actual, string? expected, string? systemDefault)
    {
        if (!hasSession) return AppRouteState.WaitingForSession;
        if (!querySucceeded) return AppRouteState.Unknown;
        if (string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase)) return AppRouteState.Matched;
        var effectiveActual = actual ?? systemDefault;
        var effectiveExpected = expected ?? systemDefault;
        if (effectiveActual == null || effectiveExpected == null) return AppRouteState.Unknown;
        return string.Equals(effectiveActual, effectiveExpected, StringComparison.OrdinalIgnoreCase)
            ? AppRouteState.Matched : AppRouteState.Drifted;
    }
}
