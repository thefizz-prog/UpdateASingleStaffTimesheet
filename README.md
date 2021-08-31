# UpdateASingleStaffTimesheet
A Power Automate Desktop flow and set of VBA subs to update staff timesheet info across all project cost control file for a given week in the past


The purpose of "UpdateASingleStaffTimesheet" is a combination of Power Automate Desktop and VBA subs to allow the office PA to update a staffs timesheet info across all project cost control files and holiday, absence and leave trackers if the staff has found out that they have incorrectly entered info on their timesheet but become aware of this at a later date. 

In order to go back and change the staff's timesheet info, all project cost control files and ad holiday, absence and leave trackers are scanned for existing entries for the given staff code. If a match is found for the given staff code in any file, the cell entries are elevated into comments  (to track changes) and the cell content is cleared. The new timesheet info for the given week for the given staff code is then updated across the project cost control files and  holiday, absence and leave trackers as specified in the staff timesheet.
