using System.Collections.Generic;

using Wox.Plugin;
using Outlook = Microsoft.Office.Interop.Outlook;

public abstract class BaseOutlookFunction {


    protected const string DefaultImage= "Images\\outlook.png"; 

    protected Outlook.Application outlookApp;

    public BaseOutlookFunction(Outlook.Application outlookApp){
        this.outlookApp = outlookApp;
    }

    public abstract List<Result> Execute(string subQuery);


}