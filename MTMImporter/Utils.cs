using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using System.Collections.ObjectModel;
using Microsoft.TeamFoundation.Framework.Common;

namespace Inflectra.SpiraTest.AddOns.MTMImporter
{
    public static class Utils
    {
        public const string ALL_TEST_PLANS = "--- All Plans ---";

        /// <summary>
        /// Test Set Status IDs
        /// </summary>
        public enum TestSetStatus
        {
            NotStarted = 1,
            InProgress = 2,
            Completed = 3,
            Blocked = 4,
            Deferred = 5
        }

        public static string SafeSubstring(this string input, int length)
        {
            if (input == null)
            {
                return "";
            }
            if (input.Length <= length)
            {
                return input;
            }
            return input.Substring(0, length);
        }


        /// <summary>
        /// Gets the list of project users in a TFS project
        /// </summary>
        /// <param name="projectCollection">TFS PROJECT COLLECTION</param>
        /// <param name="projectName">TFS PROJECT NAME</param>
        /// <returns></returns>
        public static List<TeamFoundationIdentity> ListContributors(TfsTeamProjectCollection tfsTeamProjectCollection, string projectName)
        {
            List<TeamFoundationIdentity> tfsUsers = new List<TeamFoundationIdentity>();
            IIdentityManagementService ims = (IIdentityManagementService)tfsTeamProjectCollection.GetService(typeof(IIdentityManagementService));

            // get the tfs project
            ReadOnlyCollection<CatalogNode> projectNodes = tfsTeamProjectCollection.CatalogNode.QueryChildren(new[] { CatalogResourceTypes.TeamProject }, false, CatalogQueryOptions.None);
            CatalogNode projectCatalogNode = projectNodes.FirstOrDefault(c => c.Resource.DisplayName == projectName);

            if (projectCatalogNode != null && ims != null)
            {
                TeamFoundationIdentity[] groups =  ims.ListApplicationGroups(projectName, ReadIdentityOptions.None);
                foreach (TeamFoundationIdentity group in groups)
                {
                    TeamFoundationIdentity sids = ims.ReadIdentity(IdentitySearchFactor.DisplayName, group.DisplayName, MembershipQuery.Expanded, ReadIdentityOptions.IncludeReadFromSource);
                    if (sids != null)
                    {
                        tfsUsers.AddRange(ims.ReadIdentities(sids.Members, MembershipQuery.Expanded, ReadIdentityOptions.None));
                    }
                }
            }

            //Remove any duplicates (by user-id)
            return tfsUsers.GroupBy(u => u.UniqueName).Select(u => u.First()).ToList();
        }
    }
}
