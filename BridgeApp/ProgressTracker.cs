using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OPCBridge
{
    internal class ProgressTracker: IProgress<int>
    {
        private readonly Action<int> updateProgressAction;

        public ProgressTracker(Action<int> updateProgressAction)
        {
            this.updateProgressAction = updateProgressAction;
        }

        public void Report(int value)
        {
            updateProgressAction(value);
        }
    }
}
