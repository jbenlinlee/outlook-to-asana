set asanaAPIKey to "" -- API key
set workspaceId to "" -- ID of workspace in which to create tasks
set projectId to "" -- ID of project in which to create tasks
set emailTagId to "" -- ID of tag to tag tasks

set creationNotification to "Created Asana task"
set errorNotification to "Error creating Asana task"
set appName to "Outlook to Asana"

on urlencode(theText)
	set theTextEnc to ""
	repeat with eachChar in characters of theText
		set useChar to eachChar
		set eachCharNum to ASCII number of eachChar
		if eachCharNum = 32 then
			set useChar to "+"
		else if (eachCharNum ­ 42) and (eachCharNum ­ 95) and (eachCharNum < 45 or eachCharNum > 46) and (eachCharNum < 48 or eachCharNum > 57) and (eachCharNum < 65 or eachCharNum > 90) and (eachCharNum < 97 or eachCharNum > 122) then
			set firstDig to round (eachCharNum / 16) rounding down
			set secondDig to eachCharNum mod 16
			if firstDig > 9 then
				set aNum to firstDig + 55
				set firstDig to ASCII character aNum
			end if
			if secondDig > 9 then
				set aNum to secondDig + 55
				set secondDig to ASCII character aNum
			end if
			set numHex to ("%" & (firstDig as string) & (secondDig as string)) as string
			set useChar to numHex
		end if
		set theTextEnc to theTextEnc & useChar as string
	end repeat
	return theTextEnc
end urlencode

tell application "GrowlHelperApp"
	set allNotificationsList to {creationNotification, errorNotification}
	set enabledNotificationsList to allNotificationsList
	
	register as application appName default notifications enabledNotificationsList all notifications allNotificationsList
end tell

tell application "Microsoft Outlook"
	set fwaccount to exchange account "Freewheel"
	set fwinbox to inbox of fwaccount
	
	repeat with msg in messages of fwinbox
		if todo flag of msg is not not flagged then
			set msgsubject to subject of msg as string
			
			set msgcontent to plain text content of msg as string
			set taskcontent to (characters 1 thru 80 of msgcontent as string) & "É"
			
			set msgsender to sender of msg
			-- log address of msgsender
			-- log name of msgsender
			
			set taskTitle to urlencode(msgsubject) of me
			set taskNotes to urlencode("From: " & address of msgsender) of me
			set taskNotes to taskNotes & "%0A%0A---%0A" & taskcontent
			
			-- Create Asana task
			try
				set asanaJSON to do shell script "curl -u '" & asanaAPIKey & "' https://app.asana.com/api/1.0/tasks " & Â
					"-d 'name=" & taskTitle & "' " & Â
					"-d 'workspace=" & workspaceId & "' " & Â
					"-d 'projects[0]=" & projectId & "' " & Â
					"-d 'assignee=me' " & Â
					"-d 'notes=" & taskNotes & "'"
			on error number errNum
				if (errNum is not 0) then
					tell application "GrowlHelperApp"
						notify with name errorNotification title errorNotification description msgsubject application name appName
					end tell
					
					return
				end if
			end try
			
			tell application "JSON Helper"
				set asanaRecord to read JSON from asanaJSON
				set asanaRecord to |data| of asanaRecord
			end tell
			
			-- Asana IDs can be too big for Applescript's integer type so they end up as reals
			-- This formats the IDs back to integers from AS's real representation (scientific notation)
			set taskId to do shell script "printf '%.0f' " & |id| of asanaRecord
			
			-- Add e-mail tag to task that was created
			set asanaJSON to do shell script "curl -u '" & asanaAPIKey & "' https://app.asana.com/api/1.0/tasks/" & taskId & "/addTag " & Â
				"-d 'tag=" & emailTagId & "'"
			
			-- Notify via Growl
			tell application "GrowlHelperApp"
				notify with name creationNotification title creationNotification description msgsubject application name appName
			end tell
			
			-- Unflag
			set todo flag of msg to not flagged
		end if
	end repeat
	
end tell
