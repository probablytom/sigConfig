# sigConfig
A tool for configuring and syncing email signatures between Office365 and native Office for Windows within an organisation. 

## Planned structure
The program is split into two sections: 
1. A combination of services and programs running on a server which updates office365 upon the change of the /htm file used for signatures
2. A service on each machine within the organisation which users email from, which recieves messages via .NET MessageQueues to update signatures from a shared signature file. 



#### 1. Server programs
The programs running on the server are where the configuration of the system occurs. A graphical interface is used by the administrator to either A) change settings within the sigConfig system, or B) change and update the signature at both remote Office365 locations and the native office locations. 
When the .htm signature at a configured location is altered, or when an update is pushed via the GUI, an event is triggered by a Service running on the server. This Service runs a PowerShell script to update all known users of Office365 with the new signature, and additionally sends a message to users' machines notifying them that a new signature is available. 

The service running on the server will also wake up briefly and check for changes every few hours. If a change has occured that it has not notified users' machines of, or if something with the .htm file has gone awry (for example, the configured path for the signature may no longer be a path to a file that exists), an error is raised and the service attempts to handle the situation in the most appropriate way it can. (For example, sending messages for an updated signature that it hadn't notified users' machines of.)

**TODO: Make a list of configuration options.**


#### 2. Client service
Each user's machine has instlaled on it a sigConfig service. This service operates in two ways:

1. Once every arbitrary time period, the program checks the configured path of the signature to make sure it's got the latest version. 
2. Listens out for messages from the service running on the server and updates itself whenever it is told to. 


-----------------------

## Security
User details are taken by the PowerShell script responsible for updating office365 signatures from an Excel sheet. 
To make sure that the details are kept secure, the document should be visible _only_ to the administrator running the script via Windows' permissions system. 

The email signature is kept in a visible-to-all Company Shared Folder for the organisation, and are only readable by all but Admins. All other components of the server service are admin-only files. 
