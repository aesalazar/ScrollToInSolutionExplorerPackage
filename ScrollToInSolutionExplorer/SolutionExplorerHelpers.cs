#nullable enable
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ScrollToInSolutionExplorer
{
    //https://www.codeproject.com/Articles/48990/VS-SDK-VS-Package-Find-Item-in-Solution-Explorer-W#:~:text=One%20is%20to%20use%20the%20Visual%20Studio%20native,%27Open%20%22file.cpp%22%20document%27%20item%20from%20the%20context%20menu.
    //https://github.com/microsoft/VSSDK-Extensibility-Samples/tree/c6b80a4ff7023578649d7edecc8fd6cd8a34da10
    public static class SolutionExplorerHelpers
    {
        /// <summary>
        /// Searchs for a named item in the Hierarcy and, if found, selects it.
        /// </summary>
        /// <param name="package">Reference to containing package.</param>
        /// <param name="hierarchy">Hierarchy Root.</param>
        /// <param name="nodeName">Item Name to find; null for all.</param>
        /// <returns>Tree path names, in order.</returns>
        public static async Task<IList<string>> FindItemInHierarchyAsync(
            AsyncPackage package,
            IVsHierarchy hierarchy,
            string? nodeName)
        {
            var nodes = new List<HierarchyNodeData>();
            var found = await FindItemInHierarchyAsync(
                package,
                new List<HierarchyNodeData>(),
                nodes,
                nodeName ?? string.Empty,
                hierarchy,
                VSConstants.VSITEMID_ROOT,
                0,
                true,
                false
            );

            if (found)
            {
                await SelectUIHItemAsync(package, nodes);
            }

            return nodes.Select(nd => nd.ToString()).ToList();
        }

        #region Private Methods

        /// <summary>
        /// Traversed the passed named nodes, invoking expansion and selection.
        /// </summary>
        /// <param name="package">Reference to the container package.</param>
        /// <param name="nodeData">Named nodes to search for when traversing.</param>
        private static async Task SelectUIHItemAsync(AsyncPackage package, IList<HierarchyNodeData> nodeData)
        {
            //Service is registered as DTE but must cast to DTE2
            var dte2 = await package.GetServiceAsync(typeof(DTE)) as DTE2;
            if (null == dte2)
                throw new ArgumentNullException(nameof(DTE2));

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var items = dte2.ToolWindows.SolutionExplorer.UIHierarchyItems;
            foreach (var part in nodeData.Select(nd => nd.DisplayName))
            {
                foreach (UIHierarchyItem item in items)
                {
                    if (item.Name == part)
                    {
                        item.UIHierarchyItems.Expanded = true;
                        item.Select(vsUISelectionType.vsUISelectionTypeSelect);
                        items = item.UIHierarchyItems;
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Gets the item id.
        /// </summary>
        /// <param name="pointerVarient">VARIANT holding an itemid.</param>
        /// <returns>Item Id of the concerned node</returns>
        private static uint GetItemId(object pointerVarient)
        {
            return pointerVarient switch
            {
                null => VSConstants.VSITEMID_NIL,
                int => (uint)(int)pointerVarient,
                uint => (uint)pointerVarient,
                short => (uint)(short)pointerVarient,
                ushort => (ushort)pointerVarient,
                long => (uint)(long)pointerVarient,
                _ => VSConstants.VSITEMID_NIL
            };
        }

        /// <summary>
        /// Depth-first recursive search of the hierarchy to find the path to the matching end node.
        /// </summary>
        /// <param name="package">Reference to the container package.</param>
        /// <param name="currentNodes">Collection of node names to populate.</param>
        /// <param name="targetNodeName">Name of the target node to find.</param>
        /// <param name="currentHierarchy">Root heirarchy to search on and into.</param>
        /// <param name="currentItemId">Id associated with <see cref="hierarchy"/> root.</param>
        /// <param name="currentLevel">Current level or recursion.</param>
        /// <param name="IsHeirSolutionRoot">Indicates if the current <see cref="hierarchy"/> is the root of the solution.</param>
        /// <param name="IsVisibleOnlySearch">Indicates if only visible nodes are to be searched.</param>
        /// <returns>Indicate of success.</returns>
        private static async Task<bool> FindItemInHierarchyAsync(
            AsyncPackage package,
            List<HierarchyNodeData> currentNodes,
            List<HierarchyNodeData> foundNodes,
            string targetNodeName,
            IVsHierarchy currentHierarchy,
            uint currentItemId,
            int currentLevel,
            bool IsHeirSolutionRoot,
            bool IsVisibleOnlySearch
        )
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var hierGuid = typeof(IVsHierarchy).GUID;
            var hResult = currentHierarchy.GetNestedHierarchy(currentItemId, ref hierGuid, out var nestedHeirarchyPtr, out var nestedItemId);

            //Non-zero pointer means there are child objects
            if (VSConstants.S_OK == hResult && IntPtr.Zero != nestedHeirarchyPtr)
            {
                // we are responsible to release  the refcount on the out IntPtr parameter
                var nestedHierarchy = Marshal.GetObjectForIUnknown(nestedHeirarchyPtr) as IVsHierarchy;
                Marshal.Release(nestedHeirarchyPtr);
                if (nestedHierarchy != null)
                {
                    var found = await FindItemInHierarchyAsync(package, currentNodes, foundNodes, targetNodeName, nestedHierarchy, nestedItemId, currentLevel, false, IsVisibleOnlySearch);
                    if (found)
                        return true;
                }
            }
            else
            {
                //Get the name of the root node in question here and push its name
                currentHierarchy.GetCanonicalName(currentItemId, out var cononicalName);
                currentHierarchy.GetProperty(currentItemId, (int)__VSHPROPID.VSHPROPID_Name, out var displayName);
                var nodeData = new HierarchyNodeData(cononicalName, (string)displayName);
                //Debug.WriteLine($"INFO: Walking hierarchy node: {nodeData}");
                currentNodes.Add(nodeData);

                //If we find match, terminate node enumerating
                var lastNode = currentNodes.Last();
                if (lastNode.CononicalName != null && targetNodeName
                    .EndsWith(lastNode.CononicalName, StringComparison.InvariantCultureIgnoreCase))
                    {
                    foundNodes.AddRange(currentNodes);
                    return true;
                }

                ++currentLevel;
                var siblingId = IsVisibleOnlySearch || (IsHeirSolutionRoot && currentLevel == 1)
                    ? (int)__VSHPROPID.VSHPROPID_FirstVisibleChild
                    : (int)__VSHPROPID.VSHPROPID_FirstChild;

                hResult = currentHierarchy.GetProperty(currentItemId, siblingId, out var pFirstChildId);
                ErrorHandler.ThrowOnFailure(hResult);

                if (hResult == VSConstants.S_OK)
                {
                    //using Depth first search so at each level we recurse to check if the node has any children and then look for siblings.
                    uint childId = GetItemId(pFirstChildId);

                    while (childId != VSConstants.VSITEMID_NIL)
                    {
                        var found =  await FindItemInHierarchyAsync(package, currentNodes, foundNodes, targetNodeName, currentHierarchy, childId, currentLevel, false, IsVisibleOnlySearch);
                        if (found)
                        {
                            return true;
                        }
                        
                        siblingId = IsVisibleOnlySearch || (IsHeirSolutionRoot && currentLevel == 1)
                            ? (int)__VSHPROPID.VSHPROPID_NextVisibleSibling
                            : (int)__VSHPROPID.VSHPROPID_NextSibling;

                        hResult = currentHierarchy.GetProperty(childId, siblingId, out var pNextSiblingId);
                        if (VSConstants.S_OK == hResult)
                        {
                            childId = GetItemId(pNextSiblingId);
                        }
                        else
                        {
                            ErrorHandler.ThrowOnFailure(hResult);
                            break;
                        }
                    }
                }

                if (targetNodeName.Length > 0)
                    currentNodes.RemoveAt(currentNodes.Count - 1);
            }

            //node is not found in current hierarchy
            return false;
        }

        #endregion

        #region Classes

        /// <summary>
        /// Data associated with a node in an <see cref="IVsHierarchy"/>.
        /// </summary>
        public record HierarchyNodeData
        {
            /// <summary>
            /// Creates a new instance.
            /// </summary>
            /// <param name="cononicalName">Unique name given to the node in the hierarchy used for matching.</param>
            /// <param name="displayName">String value display in the UI and used to perform navigation.</param>
            public HierarchyNodeData(string cononicalName, string displayName)
            {
                CononicalName = cononicalName;
                DisplayName = displayName;
            }

            /// <summary>
            /// Unique name given to the node in the hierarchy used for matching.
            /// </summary>
            public string CononicalName { get; }

            /// <summary>
            /// String value display in the UI and used to perform navigation.
            /// </summary>
            public string DisplayName { get; }

            /// <summary>
            /// Returns a combined Name string.
            /// </summary>
            /// <returns>Combined Name string.</returns>
            public override string ToString()
            {
                return $"{DisplayName} ({CononicalName})";
            }
        }

        #endregion
    }
}

/*
#nullable enable
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ScrollToInSolutionExplorer
{
    //https://www.codeproject.com/Articles/48990/VS-SDK-VS-Package-Find-Item-in-Solution-Explorer-W#:~:text=One%20is%20to%20use%20the%20Visual%20Studio%20native,%27Open%20%22file.cpp%22%20document%27%20item%20from%20the%20context%20menu.
    //https://github.com/microsoft/VSSDK-Extensibility-Samples/tree/c6b80a4ff7023578649d7edecc8fd6cd8a34da10
    public static class SolutionExplorerHelpers
    {
        /// <summary>
        /// Searchs for a named item in the Hierarcy and, if found, selects it.
        /// </summary>
        /// <param name="package">Reference to containing package.</param>
        /// <param name="hierarchy">Hierarchy Root.</param>
        /// <param name="nodeName">Item Name to find; null for all.</param>
        /// <returns>Tree path names, in order.</returns>
        public static async Task<IList<string>> FindItemInHierarchyAsync(
            AsyncPackage package,
            IVsHierarchy hierarchy,
            string? nodeName)
        {
            var nodeNames = new List<HierarchyNodeData>();
            await FindItemInHierarchyAsync(
                package,
                nodeNames,
                nodeName ?? string.Empty,
                hierarchy,
                VSConstants.VSITEMID_ROOT,
                0,
                true,
                false
            );

            return nodeNames
                .Select(nn => $"{nn.DisplayName} ({nn.CononicalName})")
                .ToList();
        }

        public record HierarchyNodeData
        {
            public HierarchyNodeData(string cononicalName, string displayName)
            {
                CononicalName = cononicalName;
                DisplayName = displayName;
            }

            public string CononicalName { get; }

            public string DisplayName { get; }
        }

        /// <summary>
        /// Traversed the passed named nodes, invoking expansion and selection.
        /// </summary>
        /// <param name="package">Reference to the container package.</param>
        /// <param name="nodeTreeNames">Named nodes to search for when traversing.</param>
        private static async Task SelectUIHItemAsync(AsyncPackage package, IList<HierarchyNodeData> nodeData)
        {
            //Service is registered as DTE but must cast to DTE2
            var dte2 = await package.GetServiceAsync(typeof(DTE)) as DTE2;
            if (null == dte2)
                throw new ArgumentNullException(nameof(DTE2));

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var items = dte2.ToolWindows.SolutionExplorer.UIHierarchyItems;
            foreach (var part in nodeData.Select(nd => nd.CononicalName))
            {
                foreach (UIHierarchyItem item in items)
                {
                    if (item.Name == part)
                    {
                        item.UIHierarchyItems.Expanded = true;
                        item.Select(vsUISelectionType.vsUISelectionTypeSelect);
                        items = item.UIHierarchyItems;
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Gets the item id.
        /// </summary>
        /// <param name="pvar">VARIANT holding an itemid.</param>
        /// <returns>Item Id of the concerned node</returns>
        private static uint GetItemId(object pvar)
        {
            if (pvar == null) return VSConstants.VSITEMID_NIL;
            if (pvar is int) return (uint)(int)pvar;
            if (pvar is uint) return (uint)pvar;
            if (pvar is short) return (uint)(short)pvar;
            if (pvar is ushort) return (ushort)pvar;
            if (pvar is long) return (uint)(long)pvar;
            return VSConstants.VSITEMID_NIL;
        }

        private static HierarchyNodeData? GetNodeData(IVsHierarchy hierarchy, uint itemId, __VSHPROPID property)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var hResult = hierarchy.GetProperty(itemId, (int)property, out var pDisplayName);
            hierarchy.GetCanonicalName(itemId, out var pCononcialName);
            return hResult == VSConstants.S_OK 
                ? new HierarchyNodeData(pCononcialName, (string)pDisplayName)
                : null;
        }

        /// <summary>
        /// Depth-first recursive search of the hierarchy to find the path to the matching end node.
        /// </summary>
        /// <param name="package">Reference to the container package.</param>
        /// <param name="nodeNamesFound">Collection of node names to populate.</param>
        /// <param name="nodeName">Name of the target node to find.</param>
        /// <param name="hierarchy">Root heirarchy to search on and into.</param>
        /// <param name="itemid">Id associated with <see cref="hierarchy"/> root.</param>
        /// <param name="recursionLevel">Current level or recursion.</param>
        /// <param name="IsHeirSolutionRoot">Indicates if the current <see cref="hierarchy"/> is the root of the solution.</param>
        /// <param name="IsVisibleOnlySearch">Indicates if only visible nodes are to be searched.</param>
        /// <returns>Indicate of success.</returns>
        private static async Task<bool> FindItemInHierarchyAsync(
            AsyncPackage package,
            IList<HierarchyNodeData> nodeNamesFound,
            string nodeName,
            IVsHierarchy hierarchy,
            uint itemid,
            int recursionLevel,
            bool IsHeirSolutionRoot,
            bool IsVisibleOnlySearch
        )
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var hierGuid = typeof(IVsHierarchy).GUID;
            var hResult = hierarchy.GetNestedHierarchy(itemid, ref hierGuid, out var nestedHeirarchyPtr, out var nestedItemId);

            //Non-zero pointer means there are child objects
            if (VSConstants.S_OK == hResult && IntPtr.Zero != nestedHeirarchyPtr)
            {
                // we are responsible to release the refcount on the out IntPtr parameter
                var nestedHierarchy = Marshal.GetObjectForIUnknown(nestedHeirarchyPtr) as IVsHierarchy;
                Marshal.Release(nestedHeirarchyPtr);

                //If there is something nested, continue down the tree
                if (nestedHierarchy != null)
                    await FindItemInHierarchyAsync(package, nodeNamesFound, nodeName, nestedHierarchy, nestedItemId, recursionLevel, false, IsVisibleOnlySearch);
            }
            else
            {
                //Get the data of the root node in question here and push it
                var nodeData = GetNodeData(hierarchy, itemid, __VSHPROPID.VSHPROPID_Name);
                if (nodeData is null)
                    throw new NullReferenceException($"Could not retrieve data for item {itemid}.");

                nodeNamesFound.Add(nodeData);
                Debug.WriteLine($"INFO: Walking hierarchy node: {nodeData}");

                if (nodeName.Length > 0)
                {
                    //If we find match, terminate node enumerating
                    if (string.Equals(nodeName, nodeData.CononicalName, StringComparison.InvariantCulture))
                    {
                        await SelectUIHItemAsync(package, nodeNamesFound);
                        return true;
                    }
                }

                ++recursionLevel;
                var siblingPropId = IsVisibleOnlySearch || (IsHeirSolutionRoot && recursionLevel == 1)
                    ? __VSHPROPID.VSHPROPID_FirstVisibleChild
                    : __VSHPROPID.VSHPROPID_FirstChild;

                var siblingNodeData = GetNodeData(hierarchy, itemid, siblingPropId);
                if (siblingNodeData is null)
                    throw new NullReferenceException($"Could not retrieve sibling data for item {itemid}.");

                //using Depth first search so at each level we recurse to check if the node has any children and then look for siblings.
                uint childId = GetItemId(siblingNodeData.CononicalName);
                while (childId != VSConstants.VSITEMID_NIL)
                {
                    await FindItemInHierarchyAsync(package, nodeNamesFound, nodeName, hierarchy, childId, recursionLevel, false, IsVisibleOnlySearch);
                    siblingPropId = IsVisibleOnlySearch || (IsHeirSolutionRoot && recursionLevel == 1)
                        ? __VSHPROPID.VSHPROPID_NextVisibleSibling
                        : __VSHPROPID.VSHPROPID_NextSibling;

                    siblingNodeData = GetNodeData(hierarchy, childId, siblingPropId);
                    if (siblingNodeData is null)
                        throw new NullReferenceException($"Could not retrieve child data for item {childId}.");

                    childId = GetItemId(pSiblingName);
                }

                if (nodeName.Length > 0)
                {
                    nodeNamesFound.RemoveAt(nodeNamesFound.Count - 1);
                }
            }

            //node is not found in current hierarchy
            return false;
        }
    }
}

 */