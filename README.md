# Auto MECM
GUI application for automating Microsoft Endpoint Configuration Manager (MECM) tasks.

## Description
A GUI application for carrying out Microsoft Endpoint Configuration Manager (MECM) tasks with ease. MECM was formerly known as Microsoft System Centre Configuration Manager (SCCM).

This application attempts to reduce human errors by limiting tasks that can be performed and also presents itself with an easy to use GUI.

This can also be published as an application (PA) on Citrix for users to stream. You must ensure the streaming servers have the dependencies installed for it to work.

Currently, the following tools have been added to Auto MECM:
* Add server to a collection
* Remove server from a collection
* Find a servers collection
* Set maintenance window for collections (multiple)

## Getting Started

### Dependencies
* ConfigurationManager PowerShell Module (Easiest way to install this is to install MECM)

### Configure Auto MECM
Rename the `settings.conf.example` file to `settings.conf` and add your custom values.

* `LOGLOCATION` Where Auto MECM logs will be saved to. This can be a public share and should be a public share if publishing the application on Citrix
* `SITECODE` The MECM site code. Generally this is PR1
* `SITESERVER` The MECM server
* `SCHEDULECOLLECTION` The MECM path to where servers are added to collections
* `ALLSERVERSCOLLECTION` The MECM collection to where all servers are added. This can be a Windows server collection

### Executing program
Run `Auto-MECM.ps1`

### License
[LICENSE document](./LICENSE.md)