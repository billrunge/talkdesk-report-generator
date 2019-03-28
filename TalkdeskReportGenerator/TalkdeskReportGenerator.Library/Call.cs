namespace TalkdeskReportGenerator.Library
{
    public class Call
    {
        public CallType Type { get; set; }
        public int Count { get; set; }

    }

    public enum CallType
    {
        abandoned,
        inbound,
        missed,
        outbound,
        outbound_missed,
        short_abandoned,
        voicemail
    }
}
