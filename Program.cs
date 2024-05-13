using Microsoft.Exchange.WebServices.Data;

class Program
{
    static void Main()
    {
        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);

        service.Credentials = new WebCredentials("s_msexchangegis", "xxxxxxx");
        service.Url = new Uri(@"https://webmail.brb.com.br/EWS/Exchange.asmx");
        //service.Url = new Uri(@"https://outlook.office365.com/EWS/Exchange.asmx");

        Folder inbox = Folder.Bind(service, WellKnownFolderName.Inbox);
        int numEmails = 10;
        
        ItemView view = new ItemView(numEmails);
        FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, view);

        foreach (EmailMessage email in findResults)
        {
            Console.WriteLine("Assunto: " + email.Subject);
            Console.WriteLine("Remetente: " + email.From.Address);
            Console.WriteLine("Data de recebimento: " + email.DateTimeReceived);
            try
            {
                Console.WriteLine("Corpo do email: " + email.Body);
            }
            catch
            {

            }
            Console.WriteLine("---------------------------");
        }
    }
}