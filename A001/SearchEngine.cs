using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LeadsExtractor
{
    class SearchEngine
    {
        String strEngineName;
        String strEngineURL;
        bool bIsEnabled;

        public SearchEngine(String strInputEngineName,
                            String strInputEngineURL,                
                            bool bInputIsEnabled)
        {
            strEngineName = strInputEngineName;
            strEngineURL = strInputEngineURL; 
            bIsEnabled = bInputIsEnabled;
        }

        public String getEngineName()
        {
            return this.strEngineName;
        }

        public String getEngineURL()
        {
            return this.strEngineURL;
        }

        public bool getIsEnabled()
        {
            return this.bIsEnabled;
        }

        public void setEngineName(String input_strEngineName)
        {
            strEngineName = input_strEngineName;
        }

        public void setEngineURL(String input_strEngineURL)
        {
            strEngineURL = input_strEngineURL;
        }

        public void setIsEnabled(bool input_bIsEnabled)
        {
            bIsEnabled = input_bIsEnabled;
        }

    }
}
