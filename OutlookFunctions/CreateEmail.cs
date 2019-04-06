using System;
using System.Collections.Generic;
using Wox.Plugin;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// Generate an email   
/// </summary>  
public class CreateEmail : BaseOutlookFunction
{
    public CreateEmail(Outlook.Application outlookApp) : base(outlookApp){}

    public override List<Result> Execute(string subQuery)
    {
            var result = new Result
            {
                Title = "Create new email",
                IcoPath = DefaultImage,
                Action = (Func<ActionContext, bool>)(c =>
                {
                    Outlook.MailItem mailItem = (Outlook.MailItem) outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
                    
                    mailItem.Display(true);
                    return true;
                })
            };

            return new List<Result>() { result };

    }
}