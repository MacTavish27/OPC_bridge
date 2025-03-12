using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OpcNetworkDiscovery.Services
{
    public class NetworkDiscoveryService
    {
        public event EventHandler<DiscoveryProgressEventArgs> DiscoveryProgress;

        public async Task<List<string>> DiscoverNetworkHostsAsync()
        {
            HashSet<string> hostNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            await UpdateProgressAsync("Starting network discovery...", 0);

            try
            {
                    await UpdateProgressAsync("Checking ARP cache for additional hosts...", 90);
                    await GetHostsFromArpTableAsync(hostNames);
            }
            catch (Exception ex)
            {
                await UpdateProgressAsync($"Error: {ex.Message}", -1);
                throw;
            }

            // Convert to a sorted list
            List<string> result = hostNames.OrderBy(h => h).ToList();
            await UpdateProgressAsync($"Discovery complete. Found {result.Count} hosts.", 100);
            return result;
        }

        public async Task GetHostsFromArpTableAsync(HashSet<string> hostNames)
        {
            try
            {
                Process process = new Process();
                process.StartInfo.FileName = "arp";
                process.StartInfo.Arguments = "-a";
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.CreateNoWindow = true;
                process.Start();

                string output = await process.StandardOutput.ReadToEndAsync();
                await Task.Run(() => process.WaitForExit());

                string ipPattern = @"(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})";
                MatchCollection matches = Regex.Matches(output, ipPattern);

                // Process IPs concurrently
                var tasks = matches.Cast<Match>()
                    .Where(m => m.Success)
                    .Select(m => ProcessIpAsync(m.Groups[1].Value, hostNames))
                    .ToArray();

                await Task.WhenAll(tasks);
            }
            catch (Exception ex)
            {
                await UpdateProgressAsync($"ARP discovery error: {ex.Message}", -1);
            }
        }

        public async Task ProcessIpAsync(string ipAddress, HashSet<string> hostNames)
        {
            if (ipAddress.Equals("127.0.0.1") || ipAddress.StartsWith("224.") || ipAddress.StartsWith("239."))
                return;

            try
            {
                IPHostEntry hostEntry = await Dns.GetHostEntryAsync(ipAddress);
                if (!string.IsNullOrEmpty(hostEntry.HostName) && !hostNames.Contains(hostEntry.HostName))
                {
                    string computerName = hostEntry.HostName.Split('.')[0];
                    hostNames.Add(computerName);
                    await UpdateProgressAsync($"Found host via ARP: {computerName}", -1);
                }
            }
            catch
            {
                if (!hostNames.Contains(ipAddress))
                {
                    hostNames.Add(ipAddress);
                    await UpdateProgressAsync($"Found host via ARP (IP only): {ipAddress}", -1);
                }
            }
        }

        private async Task UpdateProgressAsync(string message, int percentComplete)
        {
            // Check if there are any subscribers
            if (DiscoveryProgress != null)
            {
                // Create the event args
                var args = new DiscoveryProgressEventArgs(message, percentComplete);

                // Invoke the event handler asynchronously
                await Task.Run(() => DiscoveryProgress.Invoke(this, args));
            }
        }
    }

    public class DiscoveryProgressEventArgs : EventArgs
    {
        public string Message { get; }
        public int PercentComplete { get; }

        public DiscoveryProgressEventArgs(string message, int percentComplete)
        {
            Message = message;
            PercentComplete = percentComplete;
        }
    }
}