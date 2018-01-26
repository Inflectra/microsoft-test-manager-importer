using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Inflectra.SpiraTest.AddOns.MTMImporter
{
    /// <summary>
    /// Stores the mapping between a MTM test case and test suite
    /// </summary>
    public class MtmTestSuiteTestCaseMapping
    {
        public MtmTestSuiteTestCaseMapping()
        {
        }

        public MtmTestSuiteTestCaseMapping(int testSuiteId, int testCaseId)
        {
            this.TestSuiteId = testSuiteId;
            this.TestCaseId = testCaseId;
        }

        public int TestSuiteId { get; set; }
        public int TestCaseId { get; set; }
    }
}
