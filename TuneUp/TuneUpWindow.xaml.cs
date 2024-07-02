using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Dynamo.Extensions;
using Dynamo.Graph.Nodes;
using Dynamo.Models;
using Dynamo.Utilities;
using Dynamo.Wpf.Extensions;

namespace TuneUp
{
    /// <summary>
    /// Interaction logic for TuneUpWindow.xaml
    /// </summary>
    public partial class TuneUpWindow : Window
    {
        private ViewLoadedParams viewLoadedParams;

        private ICommandExecutive commandExecutive;

        private ViewModelCommandExecutive viewModelCommandExecutive;

        /// <summary>
        /// The unique ID for the TuneUp view extension. 
        /// Used to identify the view extension when sending recordable commands.
        /// </summary>
        private string uniqueId;

        /// <summary>
        /// Since there is no API for height offset comparing to
        /// DynamoWindow height. Define it as static for now.
        /// </summary>
        private static double sidebarHeightOffset = 200;

        /// <summary>
        /// Indicates if the TuneUp window is initializing to prevent automatic node selection.
        /// </summary>
        public bool IsInitializing { get; set; } = true;

        /// <summary>
        /// Create the TuneUp Window
        /// </summary>
        /// <param name="vlp"></param>
        public TuneUpWindow(ViewLoadedParams vlp, string id)
        {
            InitializeComponent();
            viewLoadedParams = vlp;
            // Initialize the height of the datagrid in order to make sure
            // vertical scrollbar can be displayed correctly.
            this.NodeAnalysisTable.Height = vlp.DynamoWindow.Height - sidebarHeightOffset;

            // Set the Loaded event for the DataGrid
            this.NodeAnalysisTable.Loaded += NodeAnalysisTable_Loaded;
            this.NodeAnalysisTable.DataContextChanged += NodeAnalysisTable_DataContextChanged;

            vlp.DynamoWindow.SizeChanged += DynamoWindow_SizeChanged;
            commandExecutive = vlp.CommandExecutive;
            viewModelCommandExecutive = vlp.ViewModelCommandExecutive;
            uniqueId = id;

            // Suspend the SelectionChanged event during initialization
            SuspendSelectionChanged();
        }

        /// <summary>
        /// Handles the Loaded event for the NodeAnalysisTable to ensure no item is selected 
        /// and marks initialization as complete by setting IsInitializing to false and resuming SelectionChanged handling.
        /// </summary>
        private void NodeAnalysisTable_Loaded(object sender, RoutedEventArgs e)
        {
            // Ensure no item is selected after the DataGrid is fully loaded
            NodeAnalysisTable.SelectedItem = null;
            // Initialization complete
            IsInitializing = false;
            this.ResumeSelectionChanged();
        }

        private void NodeAnalysisTable_DataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            // Clear selection when DataContext changes
            NodeAnalysisTable.SelectedItem = null;
            // Initialization complete
            IsInitializing = false;

        }

        private void DynamoWindow_SizeChanged(object sender, System.Windows.SizeChangedEventArgs e)
        {
            // Update the new height of datagrid
            this.NodeAnalysisTable.Height = e.NewSize.Height - sidebarHeightOffset;
        }

        public void SuspendSelectionChanged()
        {
            this.NodeAnalysisTable.SelectionChanged -= NodeAnalysisTable_SelectionChanged;
        }

        public void ResumeSelectionChanged()
        {
            this.NodeAnalysisTable.SelectionChanged += NodeAnalysisTable_SelectionChanged;
        }

        private void NodeAnalysisTable_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (IsInitializing) return;

            // Get NodeModel(s) that correspond to selected row(s)
            var selectedNodes = new List<NodeModel>();
            foreach (var item in e.AddedItems)
            {
                // Check NodeModel valid before actual selection
                var nodeModel = (item as ProfiledNodeViewModel).NodeModel;
                if (nodeModel != null)
                {
                    selectedNodes.Add(nodeModel);
                }
            }

            if (selectedNodes.Count() > 0)
            {
                // Select
                var command = new DynamoModel.SelectModelCommand(selectedNodes.Select(nm => nm.GUID), ModifierKeys.None);
                commandExecutive.ExecuteCommand(command, uniqueId, "TuneUp");

                // Focus on selected
                viewModelCommandExecutive.FindByIdCommand(selectedNodes.First().GUID.ToString());
            }
        }

        internal void Dispose()
        {
            viewLoadedParams.DynamoWindow.SizeChanged -= DynamoWindow_SizeChanged;
        }

        private void RecomputeGraph_Click(object sender, RoutedEventArgs e)
        {
            (NodeAnalysisTable.DataContext as TuneUpWindowViewModel).ResetProfiling();
        }
    }
}
