# Paychex Tools in PowerShell
Sample PowerShell scripts for accessing Paychex

You must have an API account with Paychex, to which you will need to supply Client ID and Client Secret's to utilize these sample scripts (with editing).  You will also need a Company Display ID in order to access a live company, rather than the sandbox company.  You will also need to edit to provide database access, for your particular database, or remove the database portions to use directly in PowerShell.
## Get-PaychexWorkers
Demonstrates accessing the Paychex Workers API to get the list of Workers in a Company, then preparing them to be merged in to an MSSQL database.
