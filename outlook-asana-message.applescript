set nl to (ASCII character 10)

tell application "Microsoft Outlook"
	-- Get first selected message
	set msgs to current messages
	set msg to item 1 of msgs
	set subj to subject of msg
	set msg_sender to sender of msg
	set msg_time to time sent of msg
	
	-- Create task message
	set taskmsg_body to "From: " & name of msg_sender & " <" & address of msg_sender & ">" & nl & "Date: " & msg_time & nl & "Subject: " & subj & nl & nl
	set taskmsg to make new outgoing message with properties {subject:subj}
	set plain text content of taskmsg to taskmsg_body
	make new to recipient at taskmsg with properties {email address:{address:"jlee@freewheel.tv"}}
	open taskmsg
end tell