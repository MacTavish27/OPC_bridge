using System;
using System.Collections.Generic;
using System.Windows;
using System.Runtime.InteropServices;
using Opc.Da;
using Opc;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Linq;
using OfficeOpenXml;
using Microsoft.Win32;

namespace OPCBridge
{
    public partial class MainWindow : Window
    {
        private Opc.Da.Server opcServer;
        private readonly Dictionary<string, Opc.Da.Subscription> tagSubscriptions = new Dictionary<string, Opc.Da.Subscription>();
        private readonly Dictionary<string, string> TagValues = new Dictionary<string, string>();
        private readonly Services services = new Services();


        public MainWindow()
        {
            InitializeComponent();
            LoadOpcServers();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void LoadOpcServers()
        {
            try
            {
              Application.Current.Dispatcher.Invoke(() =>
                {
                    opcServerList.ItemsSource = services.GetServersList();
                    if (opcServerList.Items.Count > 0)
                        opcServerList.SelectedIndex = 0;
                });
            }
            catch (COMException ex)
            {
                MessageBox.Show($"Error fetching OPC Servers: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void LoadOpcBranches(string opcServerName)
        {
            try
            {
                await ShowLoadingState();
                await services.ConnectToServerAsync(opcServerName);
                var branchNames = await services.GetBranchNamesAsync();
                await UpdateUiWithResults(branchNames);
            }
            catch (Exception ex)
            {
                ShowError(ex);
            }
            finally
            {
                await HideLoadingState();
            }
        }

        private async Task ShowLoadingState()
        {
            await Dispatcher.InvokeAsync(() =>
            {
                loadingBar.Visibility = Visibility.Visible;
                loadingBar.Value = 0;
                opcTagsList.ItemsSource = null;
                opcIDList.ItemsSource = null;
            });
        }

        private async Task UpdateUiWithResults(List<string> branchNames) => await Dispatcher.InvokeAsync(() => opcBranchesList.ItemsSource = branchNames);

        private async Task UpdateUiWithTags(List<OpcTag> tags)
        {
            await Dispatcher.InvokeAsync(() =>
            {
                opcTagsList.ItemsSource = tags.Select(t => t.Name);
                opcIDList.ItemsSource = tags.Select(t => t.ItemId);
            });
        }

        private void ShowError(Exception ex)
        {
            MessageBox.Show($"Error fetching OPC branches: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private async Task HideLoadingState() => await Dispatcher.InvokeAsync(() => loadingBar.Visibility = Visibility.Collapsed);

        public async void LoadOpcTags(string branchName)
        {
            if (opcServerList.SelectedItem == null)
            {
                MessageBox.Show("Please select an OPC server first.");
                return;
            }

            try
            {
                await ShowLoadingState();
                var tags = await services.GetTagsForBranchAsync(branchName);
                if (tags.Any())
                {
                    await UpdateUiWithTags(tags);
                }
                else
                {
                    MessageBox.Show("No available tags found in this branch.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading tags: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                await HideLoadingState();
            }
        }

        private async void SubscribeToTag(string tagName)
        {
            try
            {
                if (tagSubscriptions.ContainsKey(tagName))
                {
                    MessageBox.Show("Already subscribed to this tag!", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                string selectedOpcServer = opcServerList.SelectedItem as string;
                if (opcServer == null || !opcServer.IsConnected)
                {
                    OpcCom.Factory factory = new OpcCom.Factory();
                    opcServer = new Opc.Da.Server(factory, null);
                    await Task.Run(() => opcServer.Connect(new URL($"opcda://localhost/{selectedOpcServer}"), new ConnectData(null)));
                }

                if (!opcServer.IsConnected)
                {
                    MessageBox.Show("OPC Server is not connected!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var opcSubscriptionState = new Opc.Da.SubscriptionState
                {
                    Name = $"Subscription_{tagName}",
                    Active = true,
                    UpdateRate = 1000
                };

                var opcSubscription = (Opc.Da.Subscription)opcServer.CreateSubscription(opcSubscriptionState);
                if (opcSubscription == null)
                {
                    MessageBox.Show("Subscription was not created!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                Item item = new Item
                {
                    ItemName = tagName,
                    ClientHandle = tagName
                };

                opcSubscription.AddItems(new Opc.Da.Item[] { item });
                TagValues[tagName] = $"{tagName}: Waiting for value...";
                UpdateListBox();

                tagSubscriptions.Add(tagName, opcSubscription);
                opcSubscription.DataChanged += OnDataChange;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error subscribing to tag: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnDataChange(object subscriptionHandle, object requestHandle, Opc.Da.ItemValueResult[] values)
        {
            foreach (var item in values)
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    string tagName = item.ClientHandle.ToString();
                    TagValues[tagName] = $"{tagName}: {item.Value}";
                    UpdateListBox();
                });
            }
        }

        private void UpdateListBox()
        {
            tagValues.Items.Clear();
            foreach (var value in TagValues.Values)
            {
                tagValues.Items.Add(value);
            }
        }

        private async Task CleanupSubscriptions()
        {
            foreach (var subscription in tagSubscriptions.Values)
            {
                subscription.DataChanged -= OnDataChange;
                subscription.Dispose();
            }
            tagSubscriptions.Clear();
            TagValues.Clear();
            tagValues.Items.Clear();

            if (opcServer != null && opcServer.IsConnected)
                await Task.Run(() => opcServer.Disconnect());
        }

        private void LoadServers_Click(object sender, RoutedEventArgs e)
        {
            if (opcServerList.SelectedItem == null)
            {
                MessageBox.Show("Please select an OPC server first!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string selectedServer = opcServerList.SelectedItem.ToString();
            LoadOpcBranches(selectedServer);
        }

        private void OpcBranchesList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            opcTagsList.ItemsSource = null;
            opcIDList.ItemsSource = null;
            if (opcBranchesList.SelectedItem == null)
                return;
            string selectedBranch = opcBranchesList.SelectedItem.ToString();
            LoadOpcTags(selectedBranch);
        }

        private void LoadTagsButton_Click(object sender, RoutedEventArgs e) => OpcBranchesList_MouseDoubleClick(sender, null);

        private async void OpcServerList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            ClearSpace();
            await CleanupSubscriptions();
            await services.DisconnectAsync();
        }

        private void OpcIDsList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (opcIDList.SelectedItem != null)
            {
                string selectedTag = opcIDList.SelectedItem.ToString();
                SubscribeToTag(selectedTag);
            }
            else
            {
                MessageBox.Show("Select an available tagID!");
            }
        }

        private void ClearSpace()
        {
            opcTagsList.ItemsSource = null;
            opcIDList.ItemsSource = null;
            opcBranchesList.ItemsSource = null;
        }

        private async void ClearButton_Click(object sender, RoutedEventArgs e) => await CleanupSubscriptions();

        private async void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            await services.DisconnectAsync();
            await CleanupSubscriptions();
        }

        private void ExportToExcelButton_Click(object sender, RoutedEventArgs e)
        {
            if (opcServer == null || !opcServer.IsConnected)
            {
                MessageBox.Show("Please connect to an OPC server first!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                List<OPCTagData> tagDataList = CollectOPCData();
                if (tagDataList.Count == 0)
                {
                    MessageBox.Show("No tags subscribed to export!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    DefaultExt = "xlsx",
                    AddExtension = true,
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    FileName = $"OPC_Data_{DateTime.Now:yyyyMMdd_HHmmss}"
                };

                bool? result = saveFileDialog.ShowDialog();
                if (result == true)
                {
                    SaveToExcel(tagDataList, saveFileDialog.FileName);
                    MessageBox.Show("Data exported to Excel successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting to Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveToExcel(List<OPCTagData> tagDataList, string filePath)
        {
            using (var package = new ExcelPackage(new System.IO.FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.Add("OPC Data");
                worksheet.Cells[1, 1].Value = "OPC Server";
                worksheet.Cells[1, 2].Value = "Branch Name";
                worksheet.Cells[1, 3].Value = "Tag Name";
                worksheet.Cells[1, 4].Value = "Tag ID";
                worksheet.Cells[1, 5].Value = "Value";
                worksheet.Cells[1, 6].Value = "Timestamp";

                using (var range = worksheet.Cells[1, 1, 1, 6])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                int row = 2;
                foreach (var tag in tagDataList)
                {
                    worksheet.Cells[row, 1].Value = tag.ServerName;
                    worksheet.Cells[row, 2].Value = tag.BranchName;
                    worksheet.Cells[row, 3].Value = tag.TagName;
                    worksheet.Cells[row, 4].Value = tag.TagId;
                    worksheet.Cells[row, 5].Value = tag.Value;
                    worksheet.Cells[row, 6].Value = tag.Timestamp;
                    worksheet.Cells[row, 6].Style.Numberformat.Format = "yyyy-mm-dd hh:mm:ss";
                    row++;
                }

                worksheet.Cells.AutoFitColumns();
                package.Save();
            }
        }

        private List<OPCTagData> CollectOPCData()
        {
            var tagDataList = new List<OPCTagData>();
            string serverName = opcServerList.SelectedItem?.ToString() ?? "Unknown Server";

            foreach (var subscription in tagSubscriptions)
            {
                string tagId = subscription.Key;
                var sub = subscription.Value;

                try
                {
                    Item[] items = sub.Items;
                    if (items.Length > 0)
                    {
                        ItemValueResult[] values = sub.Read(items);
                        foreach (var value in values)
                        {
                            string tagName = value.ItemName;
                            string branchName = tagName.Contains(".") ?
                                tagName.Substring(0, tagName.LastIndexOf(".")) : "Root";

                            tagDataList.Add(new OPCTagData
                            {
                                ServerName = serverName,
                                BranchName = branchName,
                                TagName = tagName.Split('.').Last(),
                                TagId = tagId,
                                Value = value.Value?.ToString() ?? "N/A",
                                Timestamp = value.Timestamp
                            });
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error reading tag {tagId}: {ex.Message}");
                }
            }

            return tagDataList;
        }

        private class OPCTagData
        {
            public string ServerName { get; set; }
            public string BranchName { get; set; }
            public string TagName { get; set; }
            public string TagId { get; set; }
            public string Value { get; set; }
            public DateTime Timestamp { get; set; }
        }

    }
}