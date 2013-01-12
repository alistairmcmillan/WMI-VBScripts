WMI-VBScripts
=============

A few VB scripts that I've thrown together to help diagnose Windows PC faults.

PC Health Check
---------------

Gives you basic stats about RAM, Virtual Memory, Hard Drive space, etc.

Remote Process Dump
-------------------

Gives you a list of all the running processes on a remote machine including how much RAM and VM used per process.

Check Size of Queries Folder
----------------------------

Identifies which user profiles have Microsoft Queries folders and what size they are. This is to deal with a [specific Excel bug](http://alistairmcmillan.tumblr.com/post/7627301133/favourite-microsoft-bug-of-2009) we encountered.

Slow Performance Checks
-----------------------

Give you information about remote systems that might help diagnose performance issues:

- Basic stats about the PC
- Stats about RAM usage
- Stats about page file usage and whether it is system managed or not
- Scheduled tasks
- Size of temp folders
- Size of Queries folders

Check Time of Last Boot
-----------------------

Gives you the time that the remote PC last booted

Check Time on Remote PC
-----------------------

We've had issues where people can't access server applications or web-based applications because the time on the PC was skewed too far. This script lets you quickly and easily find out the time of the remote PC.

