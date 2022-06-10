using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;


namespace ContactOrganization
{
    internal class ContactGroup
    {
        public ContactGroup(Contact con, List<string> addresses)
        {
            Con = con; 
            Addresses = addresses;
        }

        public Contact Con { get; set; }
        public List<string> Addresses { get; set; }
        //public bool Combined { get; set; }

    }
    
}
