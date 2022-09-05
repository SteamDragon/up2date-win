using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management.Automation;
using Up2dateShared;

namespace Up2dateService.SetupManager
{
    public static class ChocoHelper
    {
        private const string NugetExtension = ".nupkg";
        private static Collection<string> _result;
        private static bool _chocoInstalled;
        public static ChocoPackageInstallationStatus IsPackageInstalled(Package package)
        {
            GetPackageInfo(ref package);
            if (package.ProductCode == string.Empty ||
                package.DisplayVersion == string.Empty)
                return ChocoPackageInstallationStatus.ChocoPackageInvalid;
            
            if (!IsChocoInstalled())
            {
                return ChocoPackageInstallationStatus.ChocoNotInstalled;
            }

            string resultingLineString = _result.FirstOrDefault(item => item.StartsWith(package.ProductCode));
            string[] result = null;
            if (resultingLineString != null)
            {
                result = resultingLineString.Split(' ');
            }
            // Expecting following format:
            // ProductCode Version
            if (result == null || result.Length != 2) return ChocoPackageInstallationStatus.ChocoPackageNotInstalled;
            string installedVersion = result[1];
            return installedVersion != package.DisplayVersion ?
                ChocoPackageInstallationStatus.ChocoPackageInstalledVersionDiffers :
                ChocoPackageInstallationStatus.ChocoPackageInstalled;
        }

        public static void RefreshBuffer()
        {
            using (PowerShell ps = PowerShell.Create())
            {
                const string psCommand = @"choco list -li";
                ps.AddScript(psCommand);
                _result = ps.Invoke<string>();
            }
        }

        public static bool IsChocoInstalled()
        {
            if (_chocoInstalled) return _chocoInstalled;
            using (PowerShell ps = PowerShell.Create())
            {
                ps.AddScript("choco --version");
                Collection<string> value = ps.Invoke<string>();
                _chocoInstalled = value.Count > 0;
            }

            return _chocoInstalled;
        }

        public static void InstallChocoPackage(Package package, string logDirectory, string downloadLocation,
            string externalInstallLog)
        {
            RefreshBuffer();
            // Second powershell starting is using as intermediate process to start choco process as detached 
            // Second source is used to prevent problems in downloading dependencies
            string commandType;
            ChocoPackageInstallationStatus status = IsPackageInstalled(package);
            switch (status)
            {
                case ChocoPackageInstallationStatus.ChocoPackageInstalledVersionDiffers:
                    commandType = "upgrade";
                    break;
                case ChocoPackageInstallationStatus.ChocoPackageNotInstalled:
                    commandType = "install";
                    break;
                default:
                    return;
            }

            ProcessStartInfo info = new ProcessStartInfo("choco",
                $@"{commandType} {package.ProductCode} --version {package.DisplayVersion} " +
                $@"-s ""{downloadLocation};https://community.chocolatey.org/api/v2/""" +
                @"-y --no-progress'')")
            {
                UseShellExecute = false
            };
            Process installation = Process.Start(info);
            installation?.WaitForExit();
        }

        public static void GetPackageInfo(ref Package package)
        {
            if (IsChocoInstalled() &&
                string.Equals(Path.GetExtension(package.Filepath),
                    NugetExtension,
                    StringComparison.InvariantCultureIgnoreCase))
            {
                ChocoNugetInfo nugetInfo = ChocoNugetInfo.GetInfo(package.Filepath);
                package.DisplayName = nugetInfo.Title;
                package.ProductCode = nugetInfo.Id;
                package.DisplayVersion = nugetInfo.Version;
                package.Publisher = nugetInfo.Publisher;
            }
        }
    }
}