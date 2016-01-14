using System;
using System.Text.RegularExpressions;
using Windows.Globalization.DateTimeFormatting;
using Microsoft.Office365.OutlookServices;
using o365calendar.Common;
using o365calendar.Helpers;

namespace o365calendar.Models
{
    public class EventViewModel : ViewModelBase
    {

        private string _id;
        private string _subject;
        private string _locationDisplayName;
        private bool _isNewOrDirty;
        private DateTimeOffset _start;
        private DateTimeOffset _end;
        private TimeSpan _startTime;
        private TimeSpan _endTime;
        private string _body;
        private string _attendees;
        private IEvent _serverEventData;
        private string _displayString;
        CalendarOperations _calendarOperations = new CalendarOperations();

        public string Subject
        {
            get { return _subject; }
            set
            {
                if (SetProperty(ref _subject, value))
                {
                    IsNewOrDirty = true;
                    UpdateDisplayString();
                }
            }
        }
        public string LocationName
        {
            get { return _locationDisplayName; }
            set
            {
                if (SetProperty(ref _locationDisplayName, value))
                {
                    IsNewOrDirty = true;
                    UpdateDisplayString();
                }

            }
        }
        public DateTimeOffset Start
        {
            get { return _start; }
            set
            {

                if (SetProperty(ref _start, value))
                {
                    IsNewOrDirty = true;
                    UpdateDisplayString();
                }
            }
        }
        public TimeSpan StartTime
        {
            get { return _startTime; }
            set
            {
                if (SetProperty(ref _startTime, value))
                {
                    IsNewOrDirty = true;
                    Start = Start.Date + _startTime;
                    UpdateDisplayString();
                }

            }
        }
        public DateTimeOffset End
        {
            get { return _end; }
            set
            {
                if (SetProperty(ref _end, value))
                {
                    IsNewOrDirty = true;
                    UpdateDisplayString();
                }
            }
        }
        public TimeSpan EndTime
        {
            get { return _endTime; }
            set
            {
                if (SetProperty(ref _endTime, value))
                {
                    IsNewOrDirty = true;
                    End = End.Date + _endTime;
                    UpdateDisplayString();
                }
            }
        }
        public string BodyContent
        {
            get { return _body; }
            set
            {
                if (SetProperty(ref _body, value))
                {
                    IsNewOrDirty = true;
                }
            }
        }
        public string Attendees
        {
            get { return _attendees; }
            set
            {
                if (SetProperty(ref _attendees, value))
                {
                    IsNewOrDirty = true;
                }
            }
        }

        public bool IsNewOrDirty
        {
            get
            {
                return _isNewOrDirty;
            }
            set
            {
                if (SetProperty(ref _isNewOrDirty, value) && SaveChangesCommand != null)
                {
                    UpdateDisplayString();
                    LoggingViewModel.Instance.Information = "Press the Update Event button and we'll save the changes to your calendar";
                    SaveChangesCommand.RaiseCanExecuteChanged();
                }
            }
        }

        public string DisplayString
        {
            get
            {
                return _displayString;
            }
            set
            {
                SetProperty(ref _displayString, value);
            }
        }

        private void UpdateDisplayString()
        {
            DateTimeFormatter dateFormat = new DateTimeFormatter("month.abbreviated day hour minute");

            var startDate = (Start == DateTimeOffset.MinValue) ? string.Empty : dateFormat.Format(Start);
            var endDate = (End == DateTimeOffset.MinValue) ? string.Empty : dateFormat.Format(End);

            DisplayString = String.Format("Subject: {0} Location: {1} Start: {2} End: {3}",
                    Subject,
                    LocationName,
                    startDate,
                    endDate
                    );
            DisplayString = (IsNewOrDirty) ? DisplayString + " *" : DisplayString;

        }

        public string Id
        {
            set
            {
                _id = value;
            }

            get
            {
                return _id;
            }
        }

        public bool IsNew
        {
            get
            {
                return _serverEventData == null;
            }
        }

        public void Reset()
        {
            if (!IsNew)
            {
                initialize(_serverEventData);
            }
        }


        /// <summary>
        /// Changes a calendar event.
        /// </summary>
        public RelayCommand SaveChangesCommand { get; private set; }

        private bool CanSaveChanges()
        {
            return (IsNewOrDirty);
        }

        /// <summary>
        /// Saves changes to a calendar event on the Exchange service and
        /// updates the local collection of calendar events.
        /// </summary>
        public async void ExecuteSaveChangesCommandAsync()
        {
            string operationType = string.Empty;
            try
            {
                if (!String.IsNullOrEmpty(Id))
                {
                    LoggingViewModel.Instance.Information = "Updating event ...";
                    operationType = "update";
                    //Send changes to Exchange
                    _serverEventData = await _calendarOperations.UpdateCalendarEventAsync(
                        Id,
                        LocationName,
                        BodyContent,
                        Attendees,
                        Subject,
                        Start,
                        End,
                        StartTime,
                        EndTime);
                    IsNewOrDirty = false;
                    LoggingViewModel.Instance.Information = "The event was updated in your calendar";
                }
                else
                {
                    LoggingViewModel.Instance.Information = "Adding event ...";
                    operationType = "save";
                    //Add the event
                    //Send the add request to Exchange service with new event properties
                    Id = await _calendarOperations.AddCalendarEventAsync(
                        LocationName,
                        BodyContent,
                        Attendees,
                        Subject,
                        Start,
                        End,
                        StartTime,
                        EndTime);
                    IsNewOrDirty = false;
                    LoggingViewModel.Instance.Information = "The event was added to your calendar";
                }

            }
            catch (Exception)
            {
                LoggingViewModel.Instance.Information = string.Format("We could not {0} your calendar event in your calendar", operationType);
            }
        }

        public EventViewModel(string currentUserMail)
        {
            Subject = "New Event";
            Start = DateTime.Now;
            End = DateTime.Now;
            Id = string.Empty;
            StartTime = new TimeSpan(DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
            EndTime = new TimeSpan(DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);

            Attendees = currentUserMail;

            IsNewOrDirty = true;
            SaveChangesCommand = new RelayCommand(ExecuteSaveChangesCommandAsync, CanSaveChanges);

        }


        public EventViewModel(IEvent eventData)
        {
            initialize(eventData);
        }

        private void initialize(IEvent eventData)
        {
            _serverEventData = eventData;
            string bodyContent = string.Empty;
            if (eventData.Body != null)
                bodyContent = _serverEventData.Body.Content;

            _id = _serverEventData.Id;
            _subject = _serverEventData.Subject;
            _locationDisplayName = _serverEventData.Location.DisplayName;
            _start = (DateTimeOffset)_serverEventData.Start;
            _startTime = Start.ToLocalTime().TimeOfDay;
            _end = (DateTimeOffset)_serverEventData.End;
            _endTime = End.ToLocalTime().TimeOfDay;


            //If HTML, take text. Otherwise, use content as is
            string bodyType = _serverEventData.Body.ContentType.ToString();
            if (bodyType == "HTML")
            {
                bodyContent = Regex.Replace(bodyContent, "<[^>]*>", "");
                bodyContent = Regex.Replace(bodyContent, "\n", "");
                bodyContent = Regex.Replace(bodyContent, "\r", "");
            }
            _body = bodyContent;

            _attendees = _calendarOperations.BuildAttendeeList(_serverEventData.Attendees);

            IsNewOrDirty = false;

            SaveChangesCommand = new RelayCommand(ExecuteSaveChangesCommandAsync, CanSaveChanges);
            UpdateDisplayString();
        }
    }

}
