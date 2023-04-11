# Cyclo-Scripts

Group Actions takes CSVs previously outputted from GAM and provisions those groups in M365.  It is able to:

 - Create Distribution Groups
 - Create M365 Groups
 - Delete or Purge M365 Groups
 - Update Memberships/Ownerships
 - Create Rooms

Contains Group type intelligence.
Contains PS UI

Requiring Refactoring:
 - Upgrade d-list to M365 groups
 - Gather group info


TO-DO:
 - Integrate GAM
 - Seperate functions into installable modules
 - Help files
 - Convert ugly PS UI to WPF app

Also included is Group Migrations script.  This script can
 - Migrate any type of AzureAD group from one tenant to another, including members and all group settings.
 - Migrate any type of AD group to AzureAD, including members and all group settings.
