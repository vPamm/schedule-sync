A "simple" script that will login to a Microsoft Office account, then redirect to a Kronos schedule, scrape it, then export all the shifts to seperate .ics files to be uploaded to a nextcloud or radicale calendar.

Will check if the event already exists before uploading it.

Must add your own links and everything.

Will output a log and take screenshots of the current screens in case something goes wrong (runs in a headless window)

Notify script is setup to message via telegram, need your own bot tho.

Setup script is to get it running on Linux, useful if you want to run it in a proxmox vm.
