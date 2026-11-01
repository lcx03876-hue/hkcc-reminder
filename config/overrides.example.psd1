# Timetable overrides - one-off changes from Blackboard ANNOUNCEMENTS -- EXAMPLE
# Copy to overrides.psd1 and add your own entries.
#
# Announcements never reach the calendar feed, so record them here and
# lib\build-merged.ps1 merges them into the generated calendars.
#
# ASCII only (PowerShell 5.1 cannot parse non-ASCII here without a BOM).
#
# Action = 'Exclude' : drop the matching occurrence for that date (an EXDATE
#                      is injected), leaving all other weeks untouched.
#                      Match is tested against the event SUMMARY of
#                      data\timetable.ics (starts with the course code).
#                      Optional Kind filters on DESCRIPTION
#                      (Lecture/Tutorial/Laboratory).
# Action = 'Add'     : add a one-off event (makeup class / rescheduled session).
#                      Enabled = $false keeps the record but adds nothing.
#                      AllowDuplicate = $true forces it in even when an event
#                      already exists at the same date.
@{
    Overrides = @(
        # Drop the regular Thursday lecture on one date only.
        # @{ Action = 'Exclude'; Match = 'CS101'; Date = '2027-02-01'; Kind = 'Lecture';
        #    Note = 'moved to 29 Sep 15:00-16:00' }

        # Add a make-up tutorial that is not in the weekly timetable.
        # @{ Action = 'Add'; Enabled = $true; Code = 'CS101'; Summary = 'Make-up Tutorial';
        #    Date = '2027-02-02'; Start = '16:00'; End = '18:00'; Location = 'WK-101';
        #    Description = 'Makeup tutorial per Blackboard announcement' }

        # Add a one-off exam. Set Enabled = $false to keep it as a record only.
        # @{ Action = 'Add'; Enabled = $true; Code = 'CS101'; Summary = 'Final Exam';
        #    Date = '2027-03-01'; Start = '09:00'; End = '11:00'; Location = 'WK';
        #    Description = 'Official exam timetable' }
    )
}
