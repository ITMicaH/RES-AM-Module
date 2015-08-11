# RES AM Module

Module for retreiving all sorts of information from a RES Automation Manager SQL database. It's also possible to schedule jobs using the Web API.

Visit http://itmicah.wordpress.com

Requirements:

Info only:
* RES AM database should be a SQL database.
* User account with at least Read privileges on the database.
 
Scheduling:
* A dispatcher with the Web API enabled (The Dispatcher WebAPI is disabled by default. You can enable it by enabling the global setting WebAPI state (at Infrastructure > Datastore > Settings > Global Settings > Dispatcher WebAPI settings section)
* RES AM credentials with at least rights to schedule jobs.

Run the Connect-RESAMDatabase command first.
