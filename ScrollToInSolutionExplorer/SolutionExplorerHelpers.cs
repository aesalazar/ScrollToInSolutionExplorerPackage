#nullable enable
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;

namespace ScrollToInSolutionExplorer
{
    //https://www.codeproject.com/Articles/48990/VS-SDK-VS-Package-Find-Item-in-Solution-Explorer-W#:~:text=One%20is%20to%20use%20the%20Visual%20Studio%20native,%27Open%20%22file.cpp%22%20document%27%20item%20from%20the%20context%20menu.
    //https://github.com/microsoft/VSSDK-Extensibility-Samples/tree/c6b80a4ff7023578649d7edecc8fd6cd8a34da10
    public static class SolutionExplorerHelpers
    {
        /// <summary>
        /// Searchs for a named item in the Hierarcy and, if found, selects it.
        /// </summary>
        /// <param name="visualStudioInstance">Reference to the Visual Studio Window object.</param>
        /// <param name="hierarchy">Hierarchy Root.</param>
        /// <param name="nodeName">Item Name to find; null for all.</param>
        /// <returns>Tree path names, in order.</returns>
        public static IList<string> FindItemInHierarchy(
            DTE2 visualStudioInstance,
            IVsHierarchy hierarchy,
            string? nodeName)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var nodes = new List<HierarchyNodeData>();
            var found = FindItemInHierarchy(
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
                SelectUIHItem(visualStudioInstance, nodes);

            return nodes.Select(nd => nd.ToString()).ToList();
        }

        #region Private Methods

        /// <summary>
        /// Traversed the passed named nodes, invoking expansion and selection.
        /// </summary>
        /// <param name="visualStudioInstance">Reference to the Visual Studio Window object.</param>
        /// <param name="nodeData">Named nodes to search for when traversing.</param>
        [SuppressMessage("Usage", "VSTHRD010:Invoke single-threaded types on Main thread", Justification = "Private method that can only be called from methods that have already verified access.")]
        private static void SelectUIHItem(DTE2 visualStudioInstance, IList<HierarchyNodeData> nodeData)
        {
            var items = visualStudioInstance.ToolWindows.SolutionExplorer.UIHierarchyItems;
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
        [SuppressMessage("Usage", "VSTHRD010:Invoke single-threaded types on Main thread", Justification = "Private method that can only be called from methods that have already verified access.")]
        private static bool FindItemInHierarchy(
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
                    var found = FindItemInHierarchy(currentNodes, foundNodes, targetNodeName, nestedHierarchy, nestedItemId, currentLevel, false, IsVisibleOnlySearch);
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
                currentNodes.Add(nodeData);
                //Debug.WriteLine($"INFO: Walking hierarchy node: {nodeData}");

                //If we find match, terminate node enumerating
                var lastNode = currentNodes.Last();
                if (string.Equals(targetNodeName, lastNode.CononicalName, StringComparison.OrdinalIgnoreCase))
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
                        var found = FindItemInHierarchy(currentNodes, foundNodes, targetNodeName, currentHierarchy, childId, currentLevel, false, IsVisibleOnlySearch);
                        if (found)
                            return true;

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
            public HierarchyNodeData(string? cononicalName, string? displayName)
            {
                CononicalName = cononicalName;
                DisplayName = displayName;
            }

            /// <summary>
            /// Unique name given to the node in the hierarchy used for matching.
            /// </summary>
            public string? CononicalName { get; }

            /// <summary>
            /// String value display in the UI and used to perform navigation.
            /// </summary>
            public string? DisplayName { get; }

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