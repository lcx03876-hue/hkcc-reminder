# Course names for labelling Blackboard feed events -- EXAMPLE
# Copy to courses.psd1 and replace the sample data with your own courses.
#
# The Blackboard .ics feed carries NO course information at all, so the
# course code is attached here (by UID map, or by title rules below).
#
# ASCII only on purpose (PowerShell 5.1 encoding safety).
@{
    # code -> display name
    Courses = @{
        'CS101' = 'INTRODUCTION TO PROGRAMMING'
        'EE201' = 'BASIC ELECTRICITY AND ELECTRONICS'
        'MA201' = 'ENGINEERING MATHEMATICS'
        'EG201' = 'ENGINEERING GRAPHICS AND COMPUTING'
        'ES201' = 'INTRODUCTION TO ELECTRICAL SYSTEMS'
        'PROG-01' = 'ASSOCIATE IN ENGINEERING (Programme)'
    }

    # Title rules: assign a course by matching the event title.
    # 'Like' uses wildcards (* and ?), 'Regex' uses a regular expression.
    # Rules are evaluated in order; the FIRST match wins.
    Rules = @(
        @{ Match = 'Like';  Pattern = 'Lab assignment*';   Course = 'EE201' }
        @{ Match = 'Like';  Pattern = 'Group Project*';    Course = 'CS101' }
        @{ Match = 'Regex'; Pattern = '^Quiz \d+$';        Course = 'MA201' }
    )

    # Shown when an event cannot be matched to any course.
    UnmatchedLabel = 'UNMAPPED-COURSE'

    # When true, CSV output includes a CourseLabel column.
    EmitCsv = $true
}
