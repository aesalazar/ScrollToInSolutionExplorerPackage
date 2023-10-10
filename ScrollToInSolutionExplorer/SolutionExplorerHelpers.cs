#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using Community.VisualStudio.Toolkit;

namespace ScrollToInSolutionExplorer
{
    public static class SolutionExplorerHelpers
    {
        /// <summary>
        /// Perfoms traveral on a root node of a parent-children <see cref="SolutionItem"/>.
        /// </summary>
        /// <param name="rootSolutionItem">Base solution item to traverse inclusively.</param>
        /// <param name="solutionItemNodeVistor">Prediate action to run on each node; return false to break out of the traveral.</param>
        public static void TraverseChildren(SolutionItem rootSolutionItem, Predicate<SolutionItem> solutionItemNodeVistor)
        {
            //Prep the queue
            var queue = new Queue<SolutionItem>();
            queue.Enqueue(rootSolutionItem);

            while (queue.Any())
            {
                //Get the next parent to act on the children
                var parent = queue.Dequeue();
                if (!solutionItemNodeVistor(parent))
                    break;

                //Breakout, parent has already had action applied (except top node)
                if (parent.Children == null)
                    continue;

                //Apply action to each child and push onto the stack
                foreach (var child in parent.Children)
                {
                    if (child != null)
                        queue.Enqueue(child);
                }
            }
        }
    }
}