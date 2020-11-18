# com-outlook-to-x-calendar

> Sync outlook via COM objects to {x} calendar service


## What is this?
In short, a Python script, that syncs (1 way) Outlook calendar events with other calendars.

Outlook calendar events are retrieved by interacting via COM and then writing these events to other calendar services via their API.

## OS Support
This has only been tested on the following OS's:
[x] - Windows

## Supported Calendars
Currently, the following calendar's are supported:
[x] - Google Calendar

## Scheduling
On Windows, make use of the Task Scheduler to periodically run the provided batch script `run.bat`.

