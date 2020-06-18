using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LeadsExtractor
{
    enum LeadType 
    { 
        Email =1, 
        Phone = 2, 
        Skype=3 
    };

    class Lead
    {
        Int32 nId;
        DateTime dtTimeStamp;
        LeadType nLeadType; // 1 - Email; 2 - Phone; 3 - Skype
        String strLeadContent;
        String strLeadURL;

        public Lead(LeadType nInputLeadType,
                    String strInputLeadContent,
                    String strInputLeadURL)
        {
            nLeadType = nInputLeadType;
            strLeadContent = strInputLeadContent;
            strLeadURL = strInputLeadURL;
            dtTimeStamp = DateTime.Now;
            //nIdnId++;
        }

        public Lead(Int32 nInputLeadId,
                    LeadType nInputLeadType,
                    String strInputLeadContent,
                    String strInputLeadURL)
        {
            nId = nInputLeadId;
            nLeadType = nInputLeadType;
            strLeadContent = strInputLeadContent;
            strLeadURL = strInputLeadURL;
            dtTimeStamp = DateTime.Now;            
        }

        public LeadType getLeadType()
        {
            return this.nLeadType;
        }

        public String getLeadContent()
        {
            return this.strLeadContent;
        }

        public String getLeadURL()
        {
            return this.strLeadURL;
        }

        public DateTime getLeadTimeStamp()
        {
            return this.dtTimeStamp;
        }

        public Int32 getLeadId()
        {
            return this.nId;
        }
    }
}
