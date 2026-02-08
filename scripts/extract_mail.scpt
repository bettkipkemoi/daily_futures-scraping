-- AppleScript: extract messages with specific subject from Mail
-- Writes concatenated message bodies to stdout separated by '---MSG---'
set targetSubject to "Watchlist Summary (futures)"
set separator to "---MSG---"
set resultsList to {}
try
    tell application "Mail"
        -- search all mailboxes for messages with subject that contains targetSubject
        set found to {}
        repeat with a in every account
            repeat with mb in mailboxes of a
                try
                    set ms to (messages of mb whose subject contains targetSubject)
                on error
                    set ms to {}
                end try
                repeat with m in ms
                    set end of found to m
                end repeat
            end repeat
        end repeat
        repeat with m in found
            try
                set msgContent to content of m
            on error
                set msgContent to ""
            end try
            set end of resultsList to msgContent
        end repeat
    end tell
on error errMsg number errNum
    -- return an empty result with error status text
    return "" & "ERROR:" & errNum & ": " & errMsg
end try

if (count of resultsList) is 0 then
    return ""
else
    set AppleScript's text item delimiters to separator
    set outText to resultsList as string
    set AppleScript's text item delimiters to ""
    return outText
end if
