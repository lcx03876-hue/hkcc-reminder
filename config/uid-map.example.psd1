# UID -> course code map, maintained by scripts\label-events.ps1 -- EXAMPLE
# Copy to uid-map.psd1; it can stay empty if you only use title rules.
#
# Key   : Blackboard event UID (stable across feed refreshes)
# Value : course code from config\courses.psd1
# Hand-editing is fine too. ASCII only.
@{
    # 'GradableItem-_101871@example' = 'CS101'
}
