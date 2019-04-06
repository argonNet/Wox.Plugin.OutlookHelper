using System;
using System.Collections.Generic;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Wox.Plugin.OutlookHelper
{
    public class Main : IPlugin
    {
        private Outlook.Application outlookApp;
        private Dictionary<char, BaseOutlookFunction> functions;

        public void Init(PluginInitContext context)
        {
            this.outlookApp = new Outlook.Application();

            this.functions = new Dictionary<char, BaseOutlookFunction>();
            this.functions.Add('e', new CreateEmail(outlookApp)); 
        }

        public List<Result> Query(Query query)
        {
            string workedQuery =  query.Search.Trim();
            
            return this.functions
                .Where(F => F.Key == workedQuery[0])
                .Single()
                .Value
                .Execute(workedQuery.Substring(2));

        }
    }
}
