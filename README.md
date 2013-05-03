# Introduction

This is an Applescript that creates Asana tasks from flagged Outlook messages.
Once a task is created, the associated Outlook message is unflagged.
Message subject is used for task title
Sender and first 80 chars of message are used for task description

# Requirements
- Outlook for Mac 2011
- Growl
- An Asana account with an API key


# Setup

You'll need to configure the API key, workspace, project, and tag to create
new tasks under. These are set on the four variables at the top of the
script.

Get your Asana API key in the account settings dialog

Get the workspace ID of the Asana workspace you want tasks to be created in

> curl -u '<API key>:' https://app.asana.com/api/1.0/workspaces 

Get the project ID of the Asana project you want tasks to be created in

> curl -u '<API key>:' https://app.asana.com/api/1.0/workspaces/<workspace ID>/projects

Get the ID of the tag to assign to tasks

> curl -u '<API key>:' https://app.asana.com/api/1.0/tags
