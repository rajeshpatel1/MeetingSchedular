using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;

namespace MeetingSchedularApp
{
    public class Program
    {
        private static Helper helper = new Helper();
        private static string officeURL = ConfigurationManager.AppSettings["OfficeURL"];

        private static void Main(string[] args)
        {
            try
            {
                helper.UpdateLog("Process Initiated");
                helper.UpdateLog("-----------------------------------------------------------------------");

                helper.UpdateLog("Reading Config");
                var meetings = AppConfiguration.Instance.Meetings;
                var genericCredentials = ConfigurationManager.AppSettings["ClientSettings"];
                helper.UpdateLog("-----------------------------------------------------------------------");

                helper.UpdateLog("No of Meetings Read: " + meetings.Count);
                System.Net.ServicePointManager.ServerCertificateValidationCallback =
                    ((object sender, System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                    System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors) => true);

                var i = 0;
                foreach (MeetingConfigurationElement meeting in meetings)
                {
                    i++;
                    helper.UpdateLog("Processing Meeting " + i + " Execution");
                    helper.UpdateLog("-----------------------------------------------------------------------");

                    if (meeting.IsDisabled)
                    {
                        helper.UpdateLog("Exiting ..... Meeting execution disabled");
                        helper.UpdateLog("-----------------------------------------------------------------------");

                        continue;
                    }

                    var organiserCredentials = meeting.Token.Value;
                    if (string.IsNullOrWhiteSpace(organiserCredentials))
                    {
                        organiserCredentials = genericCredentials;
                    }

                    var startTime = meeting.MeetingStartTime.Value;
                    var meetingDurationMins = Convert.ToInt32(meeting.MeetingDurationInMins.Value);
                    var advanceBookingDays = Convert.ToInt16(meeting.AdvanceBookingAddDays.Value);
                    var senderEmail = meeting.OrganiserEmailAddress.Value;
                    var toEmail = meeting.MeetingAttendees.Value;
                    var emailSubject = meeting.MeetingEmailSubject.Value;
                    var emailBodyText = meeting.MeetingEmailBody.Value;
                    var meetingLocation = meeting.MeetingRoomFullyQualifiedName.Value;
                    var responseRequired = bool.Parse(meeting.IsResponseRequested.Value);
                    var isDailySchedule = bool.Parse(meeting.IsDailySchedule.Value);
                    var customDays = meeting.ScheduledOnDays.Value;

                    CreateAndSendMeetingInvites(startTime, meetingDurationMins, advanceBookingDays, senderEmail, organiserCredentials, toEmail, emailSubject, emailBodyText
                        , meetingLocation, responseRequired, isDailySchedule, customDays);
                    helper.UpdateLog("-----------------------------------------------------------------------");
                }
            }
            catch (System.Exception ex)
            {
                helper.UpdateLog("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@");

                helper.UpdateLog("Exception: " + ex.ToString());
                helper.UpdateLog("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@");
            }
            helper.UpdateLog("Exiting Program");
            helper.UpdateLog("-----------------------------------------------------------------------");
        }

        private static void CreateAndSendMeetingInvites(string startTime
            , int meetingDurationMins
            , int advanceBookingDays
            , string organiserEmailAddress
            , string organiserCredentials
            , string toEmailsCommaList
            , string emailSubject
            , string emailBodyText
            , string meetingRoomFullyQualifiedName
            , bool responseRequested
            , bool isDailySchedule
            , string customScheduleDaysCommaList)
        {
            try
            {
                var meetingScheduleDate = DateTime.Now.Date.AddDays(advanceBookingDays).ToString("yyyyMMddT" + startTime + "ss");
                var meetingStartDateTime = DateTime.ParseExact(meetingScheduleDate, "yyyyMMddTHHmmss", System.Globalization.CultureInfo.InvariantCulture);
                var meetingEndDate = meetingStartDateTime.AddMinutes(meetingDurationMins).ToString("yyyyMMddTHHmmss");
                var meetingEndDateTime = meetingStartDateTime.AddMinutes(meetingDurationMins);
                var isScheduleAppointment = false;

                #region Logging Meeting Details in file

                helper.UpdateLog("Meeting Details");
                helper.UpdateLog("-----------------------------------------------------------------------");
                helper.UpdateLog("Meeting Organiser:        " + organiserEmailAddress);
                helper.UpdateLog("Meeting Subject:          " + emailSubject);
                helper.UpdateLog("Meeting Body Text:        " + emailBodyText);
                helper.UpdateLog("Meeting Start :           " + meetingStartDateTime);
                helper.UpdateLog("Meeting End :             " + meetingEndDateTime);
                helper.UpdateLog("Meeting Location:         " + meetingRoomFullyQualifiedName);

                foreach (var email in toEmailsCommaList.Replace(',',';').Split(new char[] { ';' }))
                {
                    helper.UpdateLog("Meeting Attendee          :" + email);
                }
                helper.UpdateLog("Meeting Response Required:" + responseRequested);
                helper.UpdateLog("Daily Schedule:           " + isDailySchedule);
                if (!isDailySchedule)
                {
                    helper.UpdateLog("Customise to be Schedule on Days (0- Sunday):" + customScheduleDaysCommaList);
                }

                helper.UpdateLog("-----------------------------------------------------------------------");

                #endregion Logging Meeting Details in file

                #region Check if scheduled to be run today

                try
                {
                    if (isDailySchedule)
                    {
                        isScheduleAppointment = true;
                    }
                    else
                    {
                        var daysToRunSchedule = customScheduleDaysCommaList.Split(',');
                        if (!string.IsNullOrEmpty(customScheduleDaysCommaList) && customScheduleDaysCommaList != "0" && daysToRunSchedule.Length >= 1)
                        {
                            foreach (string day in daysToRunSchedule)
                            {
                                if (int.Parse(day) == (int)meetingStartDateTime.DayOfWeek)
                                {
                                    isScheduleAppointment = true;
                                    break;
                                }
                            }
                        }

                        if (!isScheduleAppointment)
                        {
                            helper.UpdateLog("Custom Schedule: Meeting Scheduled will not be created today.");
                        }
                        else
                        {
                            helper.UpdateLog("Custom Schedule: Will be created today for :" + meetingStartDateTime);
                        }
                    }
                }
                catch (Exception ex)
                {
                    helper.UpdateLog("Error in Execution: Skipping Meeting Schedule");
                    helper.UpdateLog(ex.ToString());
                    return;
                }

                #endregion Check if scheduled to be run today

                if (!isScheduleAppointment)
                {
                    helper.UpdateLog("Skipping Meeting Schedule-Meeting Schedule not configured to be created today");
                    return;
                }

                if (string.Compare(meetingStartDateTime.DayOfWeek.ToString(), "sunday", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(meetingStartDateTime.DayOfWeek.ToString(), "saturday", true) == 0)
                {
                    helper.UpdateLog("Skipping Meeting Schedule-Meeting can not be schedule on Weekends");
                    return;
                }

                helper.UpdateLog("Starting Appointment Creation");

                var exchangeService = new ExchangeService
                {
                    UseDefaultCredentials = false,
                    Credentials = new WebCredentials(organiserEmailAddress, organiserCredentials),
                    Url = new System.Uri(officeURL)
                };
                NameResolution nameResolution = null;

                helper.UpdateLog("Resolving Meeting Room Name");

                var nameResolutionCollection = exchangeService.ResolveName(meetingRoomFullyQualifiedName);
                if (nameResolutionCollection.Count == 1)
                {
                    using (IEnumerator<NameResolution> enumerator = nameResolutionCollection.GetEnumerator())
                    {
                        if (enumerator.MoveNext())
                        {
                            var current = enumerator.Current;
                            nameResolution = current;
                        }
                    }
                }

                var myOptions = new AvailabilityOptions
                {
                    MeetingDuration = meetingDurationMins,
                    RequestedFreeBusyView = FreeBusyViewType.FreeBusy
                };

                var attendees = new List<AttendeeInfo>();

                attendees.Add(new AttendeeInfo()
                {
                    SmtpAddress = nameResolution.Mailbox.Address,
                    AttendeeType = MeetingAttendeeType.Required
                });

                var freeBusyResults = exchangeService.GetUserAvailability(attendees,
                                                                     new TimeWindow(meetingStartDateTime.Date, meetingStartDateTime.AddDays(1)),
                                                                         AvailabilityData.FreeBusy,
                                                                         myOptions);

                helper.UpdateLog("Creating Appointment Object");
                var appointment = new Appointment(exchangeService)
                {
                    Subject = emailSubject,
                    Body = emailBodyText,
                    Start = meetingStartDateTime,
                    End = meetingEndDateTime,
                    IsResponseRequested = responseRequested,
                    Location = meetingRoomFullyQualifiedName
                };
                if (nameResolution != null)
                {
                    appointment.RequiredAttendees.Add(new Attendee(nameResolution.Mailbox));
                }

                foreach (var email in toEmailsCommaList.Split(new char[] { ',' }))
                {
                    appointment.RequiredAttendees.Add(email);
                }
                helper.UpdateLog("Appointment Object Created ");
                helper.UpdateLog("Saving Appointment ");

                appointment.Save(SendInvitationsMode.SendToAllAndSaveCopy);
                var item = Item.Bind(exchangeService, appointment.Id, new PropertySet(new PropertyDefinitionBase[] { ItemSchema.Subject }));
                helper.UpdateLog("Appointment Saved ");
            }
            catch (Exception ex)
            {
                helper.UpdateLog("Error in Execution - Skipping record");
                helper.UpdateLog("Details: " + ex.ToString());
            }
        }
    }
}