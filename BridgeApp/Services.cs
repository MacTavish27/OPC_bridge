using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using OpcRcw.Comn; // For COM-based OPC DA
using Opc.Da;
using Opc;
using System.Threading.Tasks;
using System.Linq;

namespace OPCBridge
{

    internal class Services
    {
        private readonly OpcCom.Factory factory;
        private Opc.Da.Server opcServer;
        private bool isConnected = false;
        public Services()
        {
            factory = new OpcCom.Factory();
            opcServer = new Opc.Da.Server(factory, null);
        }

        public async Task ConnectToServerAsync(string opcServerName)
        {
            try
            {
                await Task.Run(() =>
                {
                    opcServer = new Opc.Da.Server(factory, null);
                    opcServer.Connect(new URL($"opcda://localhost/{opcServerName}"), new ConnectData(null));
                });
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to connect to OPC server: {ex.Message}");
            }
        }

        public async Task DisconnectAsync()
        {
            await Task.Run(() =>
            {
                if (opcServer != null && isConnected)
                {
                    opcServer.Disconnect();
                    isConnected = false;
                }
            });
        }
        public async Task<List<string>> GetBranchNamesAsync()
        {
            return await Task.Run(() =>
            {
                var branchesList = new List<string>();
                var filters = new BrowseFilters { BrowseFilter = browseFilter.branch };
                var branches = opcServer.Browse(null, filters, out BrowsePosition position);

                if (branches != null)
                {
                    foreach (var branch in branches)
                    {
                        if (branch.HasChildren)
                        {
                            branchesList.Add(branch.Name);
                        }
                    }
                }
                return branchesList;
            });
        }

        public async Task<List<OpcTag>> GetTagsForBranchAsync(string branchName)
        {
            return await Task.Run(() =>
            {
                var tags = new List<OpcTag>();
                var filters = new BrowseFilters { BrowseFilter = browseFilter.item };

                var browsedTags = opcServer.Browse(
                    new ItemIdentifier(branchName),
                    filters,
                    out BrowsePosition position
                );

                if (browsedTags != null)
                {
                    tags.AddRange(browsedTags.Select(tag => new OpcTag
                    {
                        Name = tag.Name,
                        ItemId = tag.ItemName
                    }));
                }

                return tags;
            });
        }

        public Opc.Da.Server GetCurrentServer()
        {
            return opcServer;
        }

        public bool IsConnected => isConnected;
        public void Disconnect() => opcServer?.Disconnect();

        public List<string> GetServersList() 
        {
            System.Type serverListType = System.Type.GetTypeFromProgID("OPC.ServerList.1");
            object serverListObject = Activator.CreateInstance(serverListType);
            IOPCServerList serverList = (IOPCServerList)serverListObject;

            // Get available OPC servers
            List<string> servers = new List<string>();
            Guid catId = new Guid("63D5F430-CFE4-11D1-B2C8-0060083BA1FB"); // OPC DA category
            serverList.EnumClassesOfCategories(1, new Guid[] { catId }, 0, null, out object enumGuidObj);

            IEnumGUID enumGuid = (IEnumGUID)enumGuidObj;
            IntPtr fetchedPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(Guid))); // Allocate memory for a GUID

            while (true)
            {
                enumGuid.Next(1, fetchedPtr, out int fetchedCount); // Get the next GUID

                if (fetchedCount == 0) // Stop if no more GUIDs
                    break;

                Guid fetchedGuid = (Guid)Marshal.PtrToStructure(fetchedPtr, typeof(Guid));
                serverList.GetClassDetails(ref fetchedGuid, out string progID, out _);
                servers.Add(progID);
            }

            Marshal.FreeCoTaskMem(fetchedPtr); // Free allocated memory

            return servers;
        }    

    }


}
