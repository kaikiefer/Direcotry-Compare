using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Direcotry_Compare
{
    class ADActions
    {

        public string[] getUserGroups(string userName)
        {
            try
            {
                int i = 0;
                // set up domain context
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

                // find user
                UserPrincipal user = UserPrincipal.FindByIdentity(ctx, userName);

                //pull underlying directory entry objects
                DirectoryEntry de = user.GetUnderlyingObject() as DirectoryEntry;

                //Get groups
                PrincipalSearchResult<Principal> groups = user.GetAuthorizationGroups();

                //Get number of groups for array creation
                int numberofGroups = groups.Count();

                //Create array for groups
                string[] groupList = new string[numberofGroups];

                // iterate over all groups
                foreach (Principal p in groups)
                {
                    // make sure to add only group principals
                    if (p is GroupPrincipal)
                    {
                        groupList[i] = p.ToString();
                        i++;
                    }
                }

                //Return the group list
                return groupList;
            }
            catch (Exception e)
            {
                //Return nothing and tell user about the error
                string[] groupList = new string[1];
                groupList[0] = "Not Found";
                MessageBox.Show("Unable to find the user: " + userName + " and the exception is: " + e);
                return groupList;
            }

        }
    }
}
