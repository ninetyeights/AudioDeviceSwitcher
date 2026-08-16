using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Text.Json;
using System.Windows;

namespace AudioDeviceSwitcher;

public record UpdateInfo(
    Version Version,
    string TagName,
    string HtmlUrl,
    string? ReleaseNotes,
    string? InstallerAssetUrl,
    string? InstallerAssetName,
    long InstallerAssetSize);

public enum UpdateCheckStatus { UpdateAvailable, UpToDate, Error }

public record UpdateCheckResult(UpdateCheckStatus Status, UpdateInfo? Update, string? ErrorMessage);

// Checks GitHub Releases for a newer build, and can download + launch the installer.
// The installer (installer/AudioDeviceSwitcher.iss) requires admin (PrivilegesRequired=admin)
// and closes/relaunches the running app itself (CloseApplications=yes), so "installing an
// update" here is just: download the asset, launch it elevated, exit this process.
public static class UpdateService
{
    // Update checks read from the org mirror, not the personal origin repo — this is the
    // canonical distribution point releases are expected to be published to.
    private const string Owner = "secure-artifacts";
    private const string Repo = "AudioDeviceSwitcher";

    private static readonly HttpClient Http = CreateClient();

    private static HttpClient CreateClient()
    {
        var client = new HttpClient { Timeout = TimeSpan.FromSeconds(15) };
        client.DefaultRequestHeaders.UserAgent.ParseAdd("AudioDeviceSwitcher-App");
        client.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
        return client;
    }

    public static Version GetCurrentVersion()
    {
        var v = Assembly.GetExecutingAssembly().GetName().Version ?? new Version(0, 0, 0);
        return Normalize(v);
    }

    // AssemblyVersion always carries 4 parts (missing ones default to 0); a GitHub tag like
    // "v1.5.0" parses to a 3-part Version with Revision == -1. Version.CompareTo treats an
    // unset (-1) component as less than a set (0) one, so without normalizing, "1.5.0" would
    // compare as older than "1.5.0.0" even though they're the same release — always compare
    // through this to sidestep that.
    private static Version Normalize(Version v) =>
        new(Math.Max(v.Major, 0), Math.Max(v.Minor, 0), Math.Max(v.Build, 0));

    public static async Task<UpdateCheckResult> CheckAsync(CancellationToken ct = default)
    {
        try
        {
            using var resp = await Http.GetAsync(
                $"https://api.github.com/repos/{Owner}/{Repo}/releases/latest", ct);
            if (!resp.IsSuccessStatusCode)
                return new UpdateCheckResult(UpdateCheckStatus.Error, null,
                    $"检查更新失败：HTTP {(int)resp.StatusCode}");

            await using var stream = await resp.Content.ReadAsStreamAsync(ct);
            using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: ct);
            var root = doc.RootElement;

            var tagName = root.TryGetProperty("tag_name", out var tagEl) ? tagEl.GetString() ?? "" : "";
            var versionText = tagName.StartsWith('v') ? tagName[1..] : tagName;
            if (!Version.TryParse(versionText, out var latestVersion))
                return new UpdateCheckResult(UpdateCheckStatus.Error, null, "无法解析发布版本号");
            latestVersion = Normalize(latestVersion);

            var htmlUrl = root.TryGetProperty("html_url", out var h) ? h.GetString() ?? "" : "";
            var body = root.TryGetProperty("body", out var b) ? b.GetString() : null;

            string? assetUrl = null, assetName = null;
            long assetSize = 0;
            if (root.TryGetProperty("assets", out var assets))
            {
                foreach (var asset in assets.EnumerateArray())
                {
                    var name = asset.TryGetProperty("name", out var n) ? n.GetString() ?? "" : "";
                    if (!name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)) continue;
                    assetUrl = asset.TryGetProperty("browser_download_url", out var u) ? u.GetString() : null;
                    assetName = name;
                    assetSize = asset.TryGetProperty("size", out var sz) ? sz.GetInt64() : 0;
                    break;
                }
            }

            var info = new UpdateInfo(latestVersion, tagName, htmlUrl, body, assetUrl, assetName, assetSize);
            return latestVersion > GetCurrentVersion()
                ? new UpdateCheckResult(UpdateCheckStatus.UpdateAvailable, info, null)
                : new UpdateCheckResult(UpdateCheckStatus.UpToDate, info, null);
        }
        catch (Exception ex)
        {
            return new UpdateCheckResult(UpdateCheckStatus.Error, null, $"检查更新失败：{ex.Message}");
        }
    }

    public static async Task<string> DownloadInstallerAsync(UpdateInfo info, IProgress<double>? progress, CancellationToken ct)
    {
        if (string.IsNullOrEmpty(info.InstallerAssetUrl))
            throw new InvalidOperationException("此版本未提供安装包文件");

        var tempPath = Path.Combine(Path.GetTempPath(),
            info.InstallerAssetName ?? $"AudioDeviceSwitcher-Setup-{info.TagName}.exe");

        using var resp = await Http.GetAsync(info.InstallerAssetUrl, HttpCompletionOption.ResponseHeadersRead, ct);
        resp.EnsureSuccessStatusCode();
        var total = resp.Content.Headers.ContentLength ?? info.InstallerAssetSize;

        await using var httpStream = await resp.Content.ReadAsStreamAsync(ct);
        await using (var fileStream = File.Create(tempPath))
        {
            var buffer = new byte[81920];
            long readTotal = 0;
            int read;
            while ((read = await httpStream.ReadAsync(buffer, ct)) > 0)
            {
                await fileStream.WriteAsync(buffer.AsMemory(0, read), ct);
                readTotal += read;
                if (total > 0) progress?.Report((double)readTotal / total);
            }
        }

        return tempPath;
    }

    public static void LaunchInstaller(string installerPath)
    {
        Process.Start(new ProcessStartInfo(installerPath)
        {
            UseShellExecute = true,
            Verb = "runas",
        });
    }

    // Shared entry point for every "check for updates" trigger (tray menu, main-window menu,
    // settings button, startup auto-check) so the UI reaction stays consistent no matter where
    // it's invoked from. manual=true always surfaces a result (including "up to date" / error);
    // manual=false (background auto-check) stays completely silent except when a genuinely new,
    // non-skipped version is found.
    public static async Task CheckAndPromptAsync(Window? owner, bool manual)
    {
        var result = await CheckAsync();

        var settings = SettingsService.Load();
        settings.LastUpdateCheckUtc = DateTime.UtcNow;

        switch (result.Status)
        {
            case UpdateCheckStatus.Error:
                SettingsService.Save();
                if (manual) ShowMessage(owner, result.ErrorMessage ?? "检查更新失败，请检查网络连接。",
                    MessageBoxImage.Warning);
                return;

            case UpdateCheckStatus.UpToDate:
                SettingsService.Save();
                if (manual) ShowMessage(owner, $"当前已是最新版本 v{GetCurrentVersion().ToString(3)}。",
                    MessageBoxImage.Information);
                return;

            case UpdateCheckStatus.UpdateAvailable:
                var info = result.Update!;
                var versionKey = info.Version.ToString(3);
                if (!manual && settings.SkippedUpdateVersion == versionKey)
                {
                    SettingsService.Save();
                    return;
                }
                SettingsService.Save();

                var win = new UpdateWindow(info) { Owner = owner };
                if (owner == null) win.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                win.Show();
                win.Activate();
                return;
        }
    }

    private static void ShowMessage(Window? owner, string text, MessageBoxImage icon)
    {
        const string caption = "检查更新";
        if (owner is { IsLoaded: true, IsVisible: true })
            MessageBox.Show(owner, text, caption, MessageBoxButton.OK, icon);
        else
            MessageBox.Show(text, caption, MessageBoxButton.OK, icon);
    }
}
