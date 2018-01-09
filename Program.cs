using System;
using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Ical.Net.Serialization;

namespace OLKMR
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            var calendar = new Ical.Net.Calendar();
            calendar.Method = CalendarMethods.Publish; // Outlook needs this property "REQUEST" will update an existing event with the same UID (Unique ID) and a newer time stamp.
            CalendarEvent e = new CalendarEvent()
            { 
                Class = "Public", 
                Organizer = new Organizer(){ CommonName = "Eric Vandekerckhove", Value = new Uri("mailto:eric.vandekerckhove@esfds.com") },
                Summary = "Summary here", 
                Created = new CalDateTime(DateTime.Now),
                Description = "Description here", 
                Start = new CalDateTime(DateTime.ParseExact("15/10/2018 15:00","dd/MM/yyyy HH:mm", null)), 
                End = new CalDateTime(DateTime.ParseExact("15/10/2018 16:00","dd/MM/yyyy HH:mm", null)),
                Sequence = 0,
                Uid = Guid.NewGuid().ToString(),
                Location = "Here"
            };
    //https://tools.ietf.org/html/rfc5545#page-25
    //"ROLE" "="
    //("CHAIR"            ; Indicates chair of the calendar entity
    // "REQ-PARTICIPANT"  ; Indicates a participant whose participation is required
    // "OPT-PARTICIPANT"  ; Indicates a participant whose participation is optional
    // "NON-PARTICIPANT"  ; Indicates a participant who is copied for information purposes only
    // x-name             ; Experimental role
    // iana-token)        ; Other IANA role
    // Default is REQ-PARTICIPANT

            e.Attendees.Add( new Attendee()
            {
                CommonName = "Veerle Serneels",
                ParticipationStatus = "REQ-PARTICIPANT",
                Rsvp = true,
                Value = new Uri("mailto:veerle.serneels@vrt.be")
            });

            Alarm alarm = new Alarm()
            {
                Action = AlarmAction.Display,
                Trigger = new Trigger(TimeSpan.FromDays(-1)),
                Summary = "Meeting in 1 day"                 
            };
            e.Alarms.Add(alarm);
            
            calendar.Events.Add(e);
            var serializer = new CalendarSerializer(new SerializationContext());
            var serializedCalendar = serializer.SerializeToString(calendar);
            Console.WriteLine(serializedCalendar);
        }
    }
}
