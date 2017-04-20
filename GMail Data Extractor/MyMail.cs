namespace GMail_Data_Extractor
{
    internal class MyMail
    {
        public string From { get; set; }
        public string Date { get; set; }
        public string ReplyTo { get; set; }
        public string Subject { get; set; }
        public string To { get; set; }
        public string TextBody { get; set; }
        public string HtmlBody { get; set; }


        public MyMail(string from, string date, string replyTo, string subject, string to, string textBody, string htmlBody)
        {
            this.From = from;
            this.Date = date;
            this.ReplyTo = replyTo;
            this.Subject = subject;
            this.To = to;
            this.TextBody = textBody;
            this.HtmlBody = htmlBody;
        }
    }
}