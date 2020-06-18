using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LeadsExtractor
{
    // class URL processed
    //
    // strURLProcessed
    // nNumOfEmails
    // nNumOfPhones
    // nNumOfSkypes

    class URLprocessed
    {
        String strURLProcessed;
        Int32 nNumOfEmails;
        Int32 nNumOfPhones;
        Int32 nNumOfSkypes;
        Int32 nDepthLevel;
        String strSearchParent;
        String strSearchHistoryChain;

        public URLprocessed(String strInputURLProcessed, 
                            Int32 nInputNumOfEmails,
                            Int32 nInputNumOfPhones,
                            Int32 nInputNumOfSkypes)
        {
            strURLProcessed = strInputURLProcessed;
            nNumOfEmails = nInputNumOfEmails;
            nNumOfPhones = nInputNumOfPhones;
            nNumOfSkypes = nInputNumOfSkypes;
            nDepthLevel = 0;
            strSearchParent = "";
            strSearchHistoryChain = "";
        }

        public URLprocessed(String strInputURLProcessed,
                            Int32 nInputNumOfEmails,
                            Int32 nInputNumOfPhones,
                            Int32 nInputNumOfSkypes,
                            Int32 nInputDepthLevel,
                            String strInputSearchParent,
                            String strInputSearchHistoryChain)
        {
            strURLProcessed = strInputURLProcessed;
            nNumOfEmails = nInputNumOfEmails;
            nNumOfPhones = nInputNumOfPhones;
            nNumOfSkypes = nInputNumOfSkypes;
            nDepthLevel = nInputDepthLevel;
            strSearchParent = strInputSearchParent;
            strSearchHistoryChain = strInputSearchHistoryChain;
        }

        public String getURLProcessed()
        {
            return this.strURLProcessed;
        }

        public Int32 getnOfEmails()
        {
            return this.nNumOfEmails;
        }

        public Int32 getnOfPhones()
        {
            return this.nNumOfPhones;
        }

        public Int32 getnOfSkypes()
        {
            return this.nNumOfSkypes;
        }

        public Int32 getnDepthLevel()
        {
            return this.nDepthLevel;
        }

        public String getSearchParent()
        {
            return this.strSearchParent;
        }

        public String getSearchHistoryChain()
        {
            return this.strSearchHistoryChain;
        }
    }
}
