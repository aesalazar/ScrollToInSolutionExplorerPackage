using System;
using System.Collections.Generic;
using System.Linq;
using Community.VisualStudio.Toolkit;

namespace ScrollToInSolutionExplorer
{
    public static class SolutionExplorerHelpers
    {
        public static void TraverseChildren(this IEnumerable<SolutionItem> solutionItems, Action<SolutionItem> filterAction)
        {
            foreach (var solutionItem in solutionItems)
            {
                //Prep the queue
                var queue = new Queue<SolutionItem>();
                queue.Enqueue(solutionItem);

                while (queue.Any())
                {
                    //Get the next parent to act on the children
                    var parent = queue.Dequeue();
                    filterAction(parent);

                    //Breakout, parent has already had action applied (except top node)
                    if (parent.Children == null)
                        continue;

                    //Apply action to each child and push onto the stack
                    foreach (var child in parent.Children)
                    {
                        //filterAction(child);
                        queue.Enqueue(child);
                    }
                }
            }
        }
    }
}
