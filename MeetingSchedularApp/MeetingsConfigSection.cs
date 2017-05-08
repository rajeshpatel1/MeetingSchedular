using System;
using System.Configuration;

namespace MeetingSchedularApp
{
    #region Config Section

    public class AppConfiguration : ConfigurationSection
    {
        private const string sectionName = "MeetingsSection";
        private const string meetingsConfigName = "Meetings";
        private const string meetingConfigName = "Meeting";

        private AppConfiguration()
        {
        }

        private static AppConfiguration instance;

        public static AppConfiguration Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = ConfigurationManager.GetSection(sectionName) as AppConfiguration;
                }

                return instance;
            }
        }

        [ConfigurationProperty(meetingsConfigName, IsRequired = true)]
        [ConfigurationCollection(typeof(MeetingConfigurationElement), AddItemName = meetingConfigName)]
        public MeetingsConfigurationsElementCollection Meetings
        {
            get
            {
                return this[meetingsConfigName] as MeetingsConfigurationsElementCollection;
            }
            set
            {
                this[meetingsConfigName] = value;
            }
        }
    }

    #endregion Config Section

    #region Meeting Collection

    public class MeetingsConfigurationsElementCollection : ConfigurationElementCollection
    {

        public MeetingConfigurationElement this[int index]
        {
            get
            {
                return (MeetingConfigurationElement)BaseGet(index);
            }
            set
            {
                if (BaseGet(index) != null)
                {
                    BaseRemoveAt(index);
                }
                BaseAdd(index, value);
            }
        }

        protected override ConfigurationElement CreateNewElement()
        {
            return new MeetingConfigurationElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return Guid.NewGuid(); //((MeetingConfigurationElement)element).MeetingStartTime;
        }
    }




    public class MeetingConfigurationElement : ConfigurationElement
    {
        private const string meetingStartTime = "MeetingStartTime";
        private const string meetingDurationInMins = "MeetingDurationInMins";
        private const string advanceBookingAddDays = "AdvanceBookingAddDays";
        private const string organiserEmailAddress = "OrganiserEmailAddress";
        private const string meetingEmailSubject = "MeetingEmailSubject";
        private const string meetingEmailBody = "MeetingEmailBody";
        private const string meetingRoomFullyQualifiedName = "MeetingRoomFullyQualifiedName";
        private const string isResponseRequested = "IsResponseRequested";
        private const string isDailySchedule = "IsDailySchedule";
        private const string scheduledOnDays = "ScheduledOnDays";

        private const string meetingAttendees = "MeetingAttendees";
        private const string token = "Token";
        [ConfigurationProperty(token, IsRequired = false)]
        public Token Token
        {
            get
            {
                var val = this[token];
                return val != null ? this[token] as Token : new Token();
            }
            set
            {
                this[token] = value;
            }
        }

        public MeetingConfigurationElement()
        {
        }

        private const string isDisabled = "isDisabled";

        [ConfigurationProperty(isDisabled, IsRequired = false, DefaultValue = false)]
        public bool IsDisabled
        {
            get
            {
                return (bool)this[isDisabled];
            }
            set
            {
                this[isDisabled] = value;
            }
        }

        [ConfigurationProperty(meetingStartTime, IsRequired = true)]
        public MeetingStartTime MeetingStartTime
        {
            get
            {
                return this[meetingStartTime] as MeetingStartTime;
            }
            set
            {
                this[meetingStartTime] = value;
            }
        }

        [ConfigurationProperty(meetingDurationInMins, IsRequired = true)]
        public MeetingDurationInMins MeetingDurationInMins
        {
            get
            {
                return this[meetingDurationInMins] as MeetingDurationInMins;
            }
            set
            {
                this[meetingDurationInMins] = value;
            }
        }

        [ConfigurationProperty(advanceBookingAddDays, IsRequired = true)]
        public AdvanceBookingAddDays AdvanceBookingAddDays
        {
            get
            {
                return this[advanceBookingAddDays] as AdvanceBookingAddDays;
            }
            set
            {
                this[advanceBookingAddDays] = value;
            }
        }

        [ConfigurationProperty(organiserEmailAddress, IsRequired = true)]
        public OrganiserEmailAddress OrganiserEmailAddress
        {
            get
            {
                return this[organiserEmailAddress] as OrganiserEmailAddress;
            }
            set
            {
                this[organiserEmailAddress] = value;
            }
        }

        [ConfigurationProperty(meetingEmailSubject, IsRequired = true)]
        public MeetingEmailSubject MeetingEmailSubject
        {
            get
            {
                return this[meetingEmailSubject] as MeetingEmailSubject;
            }
            set
            {
                this[meetingEmailSubject] = value;
            }
        }

        [ConfigurationProperty(meetingEmailBody, IsRequired = true)]
        public MeetingEmailBody MeetingEmailBody
        {
            get
            {
                return this[meetingEmailBody] as MeetingEmailBody;
            }
            set
            {
                this[meetingEmailBody] = value;
            }
        }

        [ConfigurationProperty(meetingRoomFullyQualifiedName, IsRequired = true)]
        public MeetingRoomFullyQualifiedName MeetingRoomFullyQualifiedName
        {
            get
            {
                return this[meetingRoomFullyQualifiedName] as MeetingRoomFullyQualifiedName;
            }
            set
            {
                this[meetingRoomFullyQualifiedName] = value;
            }
        }

        [ConfigurationProperty(isResponseRequested, IsRequired = true)]
        public IsResponseRequested IsResponseRequested
        {
            get
            {
                return this[isResponseRequested] as IsResponseRequested;
            }
            set
            {
                this[isResponseRequested] = value;
            }
        }

        [ConfigurationProperty(isDailySchedule, IsRequired = true)]
        public IsDailySchedule IsDailySchedule
        {
            get
            {
                return this[isDailySchedule] as IsDailySchedule;
            }
            set
            {
                this[isDailySchedule] = value;
            }
        }

        [ConfigurationProperty(scheduledOnDays, IsRequired = true)]
        public ScheduledOnDays ScheduledOnDays
        {
            get
            {
                return this[scheduledOnDays] as ScheduledOnDays;
            }
            set
            {
                this[scheduledOnDays] = value;
            }
        }

        [ConfigurationProperty(meetingAttendees, IsRequired = true)]
        public MeetingAttendees MeetingAttendees
        {
            get
            {
                return this[meetingAttendees] as MeetingAttendees;
            }
            set
            {
                this[meetingAttendees] = value;
            }
        }
    }

    #endregion Meeting Collection

    #region Meeting Element

    public class Token : ConfigurationElement
    {
        private const string value = "value";

        [ConfigurationProperty(value, IsRequired = true)]
        public string Value
        {
            get
            {
                return this[value].ToString();
            }
            set
            {
                this[value] = value;
            }
        }
    }

    public class MeetingStartTime : ConfigurationElement
    {
        private const string value = "value";

        [ConfigurationProperty(value, IsRequired = true)]
        public string Value
        {
            get
            {
                return this[value].ToString();
            }
            set
            {
                this[value] = value;
            }
        }
    }

    public class MeetingDurationInMins : ConfigurationElement
    {
        private const string value = "value";

        [ConfigurationProperty(value, IsRequired = true)]
        public string Value
        {
            get
            {
                return this[value].ToString();
            }
            set
            {
                this[value] = value;
            }
        }
    }

    public class AdvanceBookingAddDays : ConfigurationElement
    {
        private const string value = "value";

        [ConfigurationProperty(value, IsRequired = true)]
        public string Value
        {
            get
            {
                return this[value].ToString();
            }
            set
            {
                this[value] = value;
            }
        }
    }

    public class OrganiserEmailAddress : ConfigurationElement
    {
        private const string value = "value";

        [ConfigurationProperty(value, IsRequired = true)]
        public string Value
        {
            get
            {
                return this[value].ToString();
            }
            set
            {
                this[value] = value;
            }
        }
    }

    public class MeetingEmailSubject : ConfigurationElement
    {
        private const string value = "value";

        [ConfigurationProperty(value, IsRequired = true)]
        public string Value
        {
            get
            {
                return this[value].ToString();
            }
            set
            {
                this[value] = value;
            }
        }
    }

    public class MeetingEmailBody : ConfigurationElement
    {
        private const string value = "value";

        [ConfigurationProperty(value, IsRequired = true)]
        public string Value
        {
            get
            {
                return this[value].ToString();
            }
            set
            {
                this[value] = value;
            }
        }
    }

    public class MeetingRoomFullyQualifiedName : ConfigurationElement
    {
        private const string value = "value";

        [ConfigurationProperty(value, IsRequired = true)]
        public string Value
        {
            get
            {
                return this[value].ToString();
            }
            set
            {
                this[value] = value;
            }
        }
    }

    public class IsResponseRequested : ConfigurationElement
    {
        private const string value = "value";

        [ConfigurationProperty(value, IsRequired = true, DefaultValue = "true")]
        public string Value
        {
            get
            {
                return this[value].ToString();
            }
            set
            {
                this[value] = value;
            }
        }
    }

    public class IsDailySchedule : ConfigurationElement
    {
        private const string value = "value";

        [ConfigurationProperty(value, IsRequired = false, DefaultValue = "1")]
        public string Value
        {
            get
            {
                return this[value].ToString();
            }
            set
            {
                this[value] = value;
            }
        }
    }

    public class ScheduledOnDays : ConfigurationElement
    {
        private const string value = "value";

        [ConfigurationProperty(value, IsRequired = false)]
        public string Value
        {
            get
            {
                return this[value].ToString();
            }
            set
            {
                this[value] = value;
            }
        }
    }

    public class MeetingAttendees : ConfigurationElement
    {
        private const string value = "value";

        [ConfigurationProperty(value, IsRequired = true)]
        public string Value
        {
            get
            {
                return this[value].ToString();
            }
            set
            {
                this[value] = value;
            }
        }
    }

    #endregion Meeting Element
}