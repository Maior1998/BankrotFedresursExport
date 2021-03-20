using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security.Principal;
using System.ServiceModel;
using System.Text;
using AngleSharp;
using AngleSharp.Dom;
using BankruptFedresursClient;

namespace ApiTests
{
    class Program
    {
        static void Main(string[] args)
        {

            //Create a new context for evaluating webpages with the given config
            //Source to be parsed
            string source = "https://bankrot.fedresurs.ru/Messages.aspx";
            IConfiguration config = Configuration.Default.WithDefaultLoader();
            IBrowsingContext context = BrowsingContext.New(config);
            IDocument document = context.OpenAsync(source).Result;
            string cellSelector = "tr.vevent td:nth-child(3)";
            IHtmlCollection<IElement> cells = document.QuerySelectorAll(cellSelector);
            IEnumerable<string> titles = cells.Select(m => m.TextContent);
        }
    }
}
