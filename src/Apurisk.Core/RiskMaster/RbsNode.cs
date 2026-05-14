using System.Collections.Generic;

namespace Apurisk.Core.RiskMaster
{
    public sealed class RbsNode
    {
        private readonly List<RbsNode> _children;

        public RbsNode(string code, string name)
        {
            Code = code;
            Name = name;
            _children = new List<RbsNode>();
        }

        public string Code { get; private set; }
        public string Name { get; private set; }
        public IList<RbsNode> Children { get { return _children; } }
    }
}
