using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace ScrollToInSolutionExplorer
{
    //https://github.com/microsoft/VSSDK-Extensibility-Samples/blob/83759e1796f6cb9bfc509b58d7c1c0099ad5210d/ArchivedSamples/Source_Code_Control_Provider/C%23/SccProvider.cs#L1227
    public static class SccProvider
    {
        ///// <summary>
        ///// Gets the list of source controllable files in the specified project
        /////// </summary>
        //public async Task<IList<string>> GetProjectFilesAsync(IVsHierarchy hierProject, uint startItemId)
        //{
        //    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        //    var projectFiles = new List<string>();
        //    var projectItems = await GetProjectItemsAsync(hierProject, startItemId);

        //    foreach (uint itemid in projectItems)
        //    {
        //        var sccFiles = await GetNodeFilesAsync(hierProject, itemid);
        //        foreach (string file in sccFiles)
        //        {
        //            projectFiles.Add(file);
        //        }
        //    }

        //    return projectFiles;
        //}


        ///// <summary>
        ///// Returns a list of source controllable files associated with the specified node
        ///// </summary>
        //public async Task<IList<string>> GetNodeFilesAsync(IVsHierarchy hier, uint itemid)
        //{
        //    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
        //    return GetNodeFiles(hier, itemid);
        //}

        /// <summary>
        /// Gets the list of ItemIDs that are nodes in the specified project, starting with the specified item
        /// </summary>
        public static async Task<IList<uint>> GetProjectItemsAsync(IVsHierarchy pHier, uint startItemid)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            List<uint> projectNodes = new List<uint>();

            // The method does a breadth-first traversal of the project's hierarchy tree
            Queue<uint> nodesToWalk = new Queue<uint>();
            nodesToWalk.Enqueue(startItemid);

            while (nodesToWalk.Count > 0)
            {
                uint node = nodesToWalk.Dequeue();
                projectNodes.Add(node);

                await DebugWalkingNodeAsync(pHier, node);

                object property;
                if (pHier.GetProperty(node, (int)__VSHPROPID.VSHPROPID_FirstChild, out property) == VSConstants.S_OK)
                {
                    uint childnode = (uint)(int)property;
                    if (childnode == VSConstants.VSITEMID_NIL)
                    {
                        continue;
                    }

                    await DebugWalkingNodeAsync(pHier, childnode);

                    if ((pHier.GetProperty(childnode, (int)__VSHPROPID.VSHPROPID_Expandable, out property) == VSConstants.S_OK && (int)property != 0) ||
                        (pHier.GetProperty(childnode, (int)__VSHPROPID2.VSHPROPID_Container, out property) == VSConstants.S_OK && (bool)property))
                    {
                        nodesToWalk.Enqueue(childnode);
                    }
                    else
                    {
                        projectNodes.Add(childnode);
                    }

                    while (pHier.GetProperty(childnode, (int)__VSHPROPID.VSHPROPID_NextSibling, out property) == VSConstants.S_OK)
                    {
                        childnode = (uint)(int)property;
                        if (childnode == VSConstants.VSITEMID_NIL)
                        {
                            break;
                        }

                        await DebugWalkingNodeAsync(pHier, childnode);

                        if ((pHier.GetProperty(childnode, (int)__VSHPROPID.VSHPROPID_Expandable, out property) == VSConstants.S_OK && (int)property != 0) ||
                            (pHier.GetProperty(childnode, (int)__VSHPROPID2.VSHPROPID_Container, out property) == VSConstants.S_OK && (bool)property))
                        {
                            nodesToWalk.Enqueue(childnode);
                        }
                        else
                        {
                            projectNodes.Add(childnode);
                        }
                    }
                }
            }

            return projectNodes;
        }

        private static async Task DebugWalkingNodeAsync(IVsHierarchy pHier, uint itemid)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            if (pHier.GetProperty(itemid, (int)__VSHPROPID.VSHPROPID_Name, out var property) == VSConstants.S_OK)
            {
                Debug.WriteLine(string.Format(CultureInfo.CurrentUICulture, "Walking hierarchy node: {0}", (string)property));
            }
        }

    }
}
