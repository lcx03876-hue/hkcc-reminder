# Deadlines that are NOT in the Blackboard calendar feed -- EXAMPLE
# Copy to extra-deadlines.psd1 and add your own entries.
#
# Why this file exists: build-merged.ps1 derives the deadline calendar from
# data\blackboard.ics, which is rewritten on every sync. Items that never
# reach the feed (assignments in Course Content, announcements) are lost if
# you edit the .ics by hand, so register them here instead.
#
# Fields:
#   Enabled     = $true/$false   (false keeps the record but adds nothing)
#   Code        = course code, must exist in config\courses.psd1 for labelling
#   Summary     = short title shown after the course code
#   Due         = 'yyyy-MM-dd HH:mm' in Asia/Hong_Kong (local)
#   Description = optional free text (ASCII)
#
# The UID is derived from code + due time, so re-importing updates the
# existing event instead of creating a duplicate.
@{
    Deadlines = @(
        # @{ Enabled = $true; Code = 'CS101'; Summary = 'Assignment 1';
        #    Key = 'CS101-A1';
        #    Due = '2027-01-15 23:59';
        #    Description = 'Submit the source code to Blackboard.' }
    )
}
