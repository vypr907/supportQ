Attribute VB_Name = "queueActions"
'For all things pertaining to the operation of a functional queue

'Proposed:
' queue is one sheet. "my queue" is filtered based off technician initials in
' respective column. "admin queue" shows all current entries. "main queue" shows
' only un-taken entries. log is merely a backup of the queue.
'
'Original Idea:
' main queue is one sheet. "my queue" is separate sheet, populated by the 'Take'
' button. Main queue grows and shrinks as users submit and technicians take.
' Upon technician RESOLVE, entry is then copied to "resolved queue" or Log sheet
' and removed from "my queue". "Admin queue" is just Log.

'TODO: loadQueue function. accepts int variable to determine which q to load
' 1. Main 2. User 3. Log

