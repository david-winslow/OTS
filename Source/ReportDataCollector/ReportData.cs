namespace ReportDataCollector
{
    public class ReportData
    {
        public Client Client { get; set; }
        public Background Background { get; set; }
        public Occupation Occupation { get; set; }

        public ReportData(Client client, Background background, Occupation occupation)
        {
            Client = client;
            Background = background;
            Occupation = occupation;
        }
    }
}