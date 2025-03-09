using Email_Reading_and_xml_writing;


var mails = OutlookEmails.ReadMailItems();
int i = 1;
foreach (var mail in mails)
{
    Console.WriteLine("Mail No " + i);
    Console.WriteLine("Mail Recieved from " + mail.EmailFrom);
    Console.WriteLine("Mail Subject " + mail.EmailSubject);
    Console.WriteLine("Mail Body " + mail.EmailBody);
    Console.WriteLine("");
    i++;
}
Console.ReadKey();