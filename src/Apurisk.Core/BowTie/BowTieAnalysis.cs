using System.Collections.Generic;
using Apurisk.Core.RiskMaster;

namespace Apurisk.Core.BowTie
{
    public sealed class BowTieAnalysis
    {
        private readonly List<Threat> _threats;
        private readonly List<Barrier> _preventiveBarriers;
        private readonly List<Consequence> _consequences;
        private readonly List<Barrier> _mitigatingBarriers;

        public BowTieAnalysis(RiskItem risk)
        {
            Risk = risk;
            Hazard = new Hazard();
            TopEvent = string.Empty;
            _threats = new List<Threat>();
            _preventiveBarriers = new List<Barrier>();
            _consequences = new List<Consequence>();
            _mitigatingBarriers = new List<Barrier>();
        }

        public RiskItem Risk { get; private set; }
        public Hazard Hazard { get; private set; }
        public string TopEvent { get; set; }
        public IList<Threat> Threats { get { return _threats; } }
        public IList<Barrier> PreventiveBarriers { get { return _preventiveBarriers; } }
        public IList<Consequence> Consequences { get { return _consequences; } }
        public IList<Barrier> MitigatingBarriers { get { return _mitigatingBarriers; } }
    }
}
