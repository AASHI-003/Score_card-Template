# Score_card-Template
Structured report on how to develop an Excel Scorecard template that uses macros (VBA) to protect specific columns so that when anyone tries to unhide them, they’re prompted for a password. This approach goes beyond standard worksheet protection and gives you controlled column-level security with a guided user experience.


Objectives

Provide a secure Scorecard template where select columns are accessible only to authorized users.
Improve usability with guided buttons for Hide/Unhide actions.
Ensure protection persists even if the workbook is re-opened.

In Scope

Column-level protection with password prompt.
Button-driven hide/unhide workflows.
Setup instructions and testing checklist.

Out of Scope

Enterprise-level encryption (this is Excel/VBA, not IRM or sensitivity labels).
Role-based access tied to Azure AD or external systems.


#Summary
This approach gives you practical, controlled access to sensitive columns in a Scorecard. Given your context, it sounds like you already figured out how to protect a worksheet—this solution extends that capability to require a password when unhiding columns in a user-friendly way.
