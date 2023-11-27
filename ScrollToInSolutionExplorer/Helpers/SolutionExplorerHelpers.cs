#nullable enable
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using VSID = Microsoft.VisualStudio.Shell.Interop.__VSHPROPID;

namespace ScrollToInSolutionExplorer
{
    /// <summary>
    /// Helper methods to navigate the Solution Explorer.
    /// </summary>
    /// <remarks>
    /// Based on posts:
    /// https://www.codeproject.com/Articles/48990/VS-SDK-VS-Package-Find-Item-in-Solution-Explorer-W#:~:text=One%20is%20to%20use%20the%20Visual%20Studio%20native,%27Open%20%22file.cpp%22%20document%27%20item%20from%20the%20context%20menu.
    /// https://github.com/microsoft/VSSDK-Extensibility-Samples/tree/c6b80a4ff7023578649d7edecc8fd6cd8a34da10
    /// </remarks>
    internal static class SolutionExplorerHelpers
    {
        /// <summary>
        /// Searches for a named item in the Hierarcy and returns an indication if found.
        /// </summary>
        /// <param name="hierarchy">Hierarchy Root.</param>
        /// <param name="displayName">Item Name to find.</param>
        /// <returns>Tree path names, in order.</returns>
        public static bool IsDisplayNameInHierarcy(
            IVsHierarchy hierarchy,
            string displayName)
        {
            if (displayName.Length == 0)
                throw new ArgumentException("Display Name cannot be an empty string.", nameof(displayName));

            ThreadHelper.ThrowIfNotOnUIThread();
            return IsDisplayNameInHierarcy(
                displayName,
                hierarchy,
                VSConstants.VSITEMID_ROOT,
                0,
                new List<HierarchyNodeData>());
        }

        #region Private Methods

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

        [SuppressMessage("Usage", "VSTHRD010:Invoke single-threaded types on Main thread", Justification = "Private method that can only be called from methods that have already verified access.")]
        private static bool IsDisplayNameInHierarcy(
            string targetDisplayName,
            IVsHierarchy currentHierarchy,
            uint currentItemId,
            int currentLevel,
            List<HierarchyNodeData> currentNodes)
        {
            var hierGuid = typeof(IVsHierarchy).GUID;
            var hResult = currentHierarchy.GetNestedHierarchy(currentItemId, ref hierGuid, out var nestedHeirarchyPtr, out var nestedItemId);

            //Non-zero pointer means there are child objects
            if (hResult == VSConstants.S_OK && nestedHeirarchyPtr != IntPtr.Zero)
            {
                if (Marshal.GetObjectForIUnknown(nestedHeirarchyPtr) is IVsHierarchy nestedHierarchy &&
                    IsDisplayNameInHierarcy(targetDisplayName, nestedHierarchy, nestedItemId, currentLevel, currentNodes))
                    return true;
            }
            else
            {
                //Get the name of the root node in question here and push its name
                currentHierarchy.GetCanonicalName(currentItemId, out var cononicalName);
                currentHierarchy.GetProperty(currentItemId, (int)VSID.VSHPROPID_Name, out var displayName);
                var nodeData = new HierarchyNodeData(cononicalName, (string)displayName);
                currentNodes.Add(nodeData);
                //Debug.WriteLine($"INFO: {nodeData}");

                //If we find match, terminate node enumerating
                var lastNode = currentNodes.Last();
                if (string.Equals(targetDisplayName, lastNode.DisplayName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                var isRootNode = ++currentLevel == 1;
                var siblingId = isRootNode ? VSID.VSHPROPID_FirstVisibleChild : VSID.VSHPROPID_FirstChild;
                hResult = currentHierarchy.GetProperty(currentItemId, (int)siblingId, out var pFirstChildId);
                ErrorHandler.ThrowOnFailure(hResult);

                if (hResult == VSConstants.S_OK)
                {
                    //using Depth first search so at each level we recurse to check if the node has any children and then look for siblings.
                    var childId = GetItemId(pFirstChildId);
                    while (childId != VSConstants.VSITEMID_NIL)
                    {
                        var found = IsDisplayNameInHierarcy(targetDisplayName, currentHierarchy, childId, currentLevel, currentNodes);
                        if (found)
                            return true;

                        siblingId = isRootNode ? VSID.VSHPROPID_NextVisibleSibling : VSID.VSHPROPID_NextSibling;
                        hResult = currentHierarchy.GetProperty(childId, (int)siblingId, out var pNextSiblingId);

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
        private record HierarchyNodeData
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