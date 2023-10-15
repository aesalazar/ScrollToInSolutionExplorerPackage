#nullable enable
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Threading;
using Microsoft.VisualStudio.Shell;
using EnvDTE;
using EnvDTE80;
using System.Diagnostics;
using System.Globalization;

namespace ScrollToInSolutionExplorer
{
    //https://www.codeproject.com/Articles/48990/VS-SDK-VS-Package-Find-Item-in-Solution-Explorer-W#:~:text=One%20is%20to%20use%20the%20Visual%20Studio%20native,%27Open%20%22file.cpp%22%20document%27%20item%20from%20the%20context%20menu.
    //https://github.com/microsoft/VSSDK-Extensibility-Samples/tree/c6b80a4ff7023578649d7edecc8fd6cd8a34da10
    public class SolutionExplorerHelpers
    {


        public SolutionExplorerHelpers(AsyncPackage package)
        {
            _package = package;
        }

        ///// <summary>
        ///// Manages work of workers
        ///// </summary>
        //public class FindNextManager
        //{
        //    /// <summary>
        //    /// Indicates that previous worker thread has to terminate its work
        //    /// </summary>
        //    public Boolean IsTermination = false;
        //    /// <summary>
        //    /// Worker thread, here work runs
        //    /// </summary>
        //    public System.Threading.Thread Worker = null;
        //}

        private async Task SelectUIHItemAsync(List<string> nodesTree)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(_package.DisposalToken);
            var dte = await _package.GetServiceAsync(typeof(DTE)) as DTE2;
            var dte2 = dte as DTE2;

            if (null == dte2)
                return;

            //UIHierarchy UIH = dte2.ToolWindows.SolutionExplorer;
            //UIHierarchyItem UIHItem;// = UIH.GetItem(path);
            //StringBuilder path = new StringBuilder();

            //for (int i = 0; i < nodesTree.Count; ++i)
            //{
            //    if (i > 0)
            //    {
            //        path.Append('\\');
            //    }
            //    path.Append(nodesTree[i]);

            //    UIHItem = UIH.GetItem(path.ToString());
            //    UIHItem.Select(vsUISelectionType.vsUISelectionTypeSelect);

            //    if (!UIHItem.UIHierarchyItems.Expanded)
            //    {
            //        UIH.DoDefaultAction();
            //    }
            //}

            var items = dte2.ToolWindows.SolutionExplorer.UIHierarchyItems;
            foreach (var part in nodesTree)
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
        private uint GetItemId(object pvar)
        {
            if (pvar == null) return VSConstants.VSITEMID_NIL;
            if (pvar is int) return (uint)(int)pvar;
            if (pvar is uint) return (uint)pvar;
            if (pvar is short) return (uint)(short)pvar;
            if (pvar is ushort) return (ushort)pvar;
            if (pvar is long) return (uint)(long)pvar;
            return VSConstants.VSITEMID_NIL;
        }

        /// <summary>
        /// Worker manager
        /// </summary>
        //private FindNextManager workerManager = new FindNextManager();
        private readonly AsyncPackage _package;

        public async Task<bool> FindItemInHierarchyAsync(
            List<string> clew, 
            string match,
            IVsHierarchy hierarchy, 
            uint itemid, 
            int recursionLevel,
            bool hierIsSolution, 
            bool visibleNodesOnly
        )
        {
            ////bool isTermination;
            ////lock (workerManager)
            ////{
            ////    isTermination = workerManager.IsTermination;
            ////}
            ////if (isTermination)
            ////{
            ////    return false;
            ////}

            int hr;
            IntPtr nestedHierarchyObj;
            uint nestedItemId;
            Guid hierGuid = typeof(IVsHierarchy).GUID;

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            //await DebugWalkingNodeAsync(hierarchy, itemid);

            hr = hierarchy.GetNestedHierarchy(itemid, ref hierGuid,
                           out nestedHierarchyObj, out nestedItemId);
            if (VSConstants.S_OK == hr && IntPtr.Zero != nestedHierarchyObj)
            {
                var nestedHierarchy = Marshal.GetObjectForIUnknown(nestedHierarchyObj) as IVsHierarchy;

                // we are responsible to release
                // the refcount on the out IntPtr parameter
                Marshal.Release(nestedHierarchyObj);
                if (nestedHierarchy != null)
                {
                    // Display name and type of the node in the Output Window
                    //if (
                    await FindItemInHierarchyAsync(
                        clew, match, nestedHierarchy, nestedItemId,
                            recursionLevel, false, visibleNodesOnly);
                    ////)
                    ////{
                    ////return true;
                    ////}

                    //if (workerManager.IsTermination)
                    //{
                    //    return false;
                    //}
                }
            }
            else
            {
                object pVar;

                //Get the name of the root node in question
                //here and push its value to clew
                hr = hierarchy.GetProperty(itemid,
                        (int)__VSHPROPID.VSHPROPID_Name, out pVar);

                clew.Add((string)pVar);

                if (match.Length > 0)
                {
                    //If we find match, terminate node enumerating
                    if (Regex.Match(match, clew[clew.Count - 1],
                          RegexOptions.IgnoreCase).Value != string.Empty)
                    {
                        await SelectUIHItemAsync(clew);

                        ////if (workerManager.IsTermination)
                        ////{
                        ////    return false;
                        ////}

                        return true;
                    }
                }
                else
                {
                    // suppose this is just request for all items in solution
                }

                ++recursionLevel;

                hr = hierarchy.GetProperty(itemid,
                    ((visibleNodesOnly || (hierIsSolution && recursionLevel == 1) ?
                        (int)__VSHPROPID.VSHPROPID_FirstVisibleChild :
                        (int)__VSHPROPID.VSHPROPID_FirstChild)), out pVar);
                Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(hr);
                if (VSConstants.S_OK == hr)
                {
                    //We are using Depth first search so at each level
                    //we recurse to check if the node has any children
                    //and then look for siblings.
                    uint childId = GetItemId(pVar);
                    while (childId != VSConstants.VSITEMID_NIL)
                    {
                        //if (
                        await FindItemInHierarchyAsync(
                            clew, match, hierarchy, childId,
                            recursionLevel, false, visibleNodesOnly);
                        ////)
                        ////{
                        ////return true
                        ////}

                        //if (workerManager.IsTermination)
                        //{
                        //    return false;
                        //}

                        hr = hierarchy.GetProperty(childId,
                            ((visibleNodesOnly || (hierIsSolution && recursionLevel == 1)) ?
                                (int)__VSHPROPID.VSHPROPID_NextVisibleSibling :
                                (int)__VSHPROPID.VSHPROPID_NextSibling),
                            out pVar);
                        if (VSConstants.S_OK == hr)
                        {
                            childId = GetItemId(pVar);
                        }
                        else
                        {
                            Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(hr);
                            break;
                        }
                    }
                }

                if (match.Length > 0)
                {
                    clew.RemoveAt(clew.Count - 1);
                }
            }

            return false; //node is not found in current hierarchy
        }

        private static async Task DebugWalkingNodeAsync(IVsHierarchy pHier, uint itemid)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            if (pHier.GetProperty(itemid, (int)__VSHPROPID.VSHPROPID_Name, out var property) == VSConstants.S_OK)
            {
                Debug.WriteLine(string.Format(CultureInfo.CurrentUICulture, "INFO: Walking hierarchy node: {0}", (string)property));
            }
        }
    }
}
