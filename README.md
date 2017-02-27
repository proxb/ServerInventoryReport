# ServerInventoryReport
A set of scripts to create an inventory report using PowerShell, SQL and WPF

![alt tag](https://github.com/proxb/ServerInventoryReport/blob/master/Images/ServerInventoryUI.png)

## Requirements for use
* SQL Server
* PowerShell V3
* Necessary rights to pull information from remote servers
* Necessary rights to run a scheduled task for data gathering
* [PoshRSJob](https://github.com/proxb/PoshRSJob) module to assist with data gathering

## Using the Server Inventory scripts (work in progress)
1. Either on the SQL server or on a system that has access to the SQL server where the database and tables will reside on, run the [ServerInventory_SQLBuild.ps1](https://github.com/proxb/ServerInventoryReport/blob/master/ServerInventory_SQLBuild.ps1) script, ensuring that you have updated the SQL database location (Computername parameter on line 4).

2. You will need to provide the account that is being used in the scheduled task the necessary rights to be able to write to the ServerInventory database that is created.

3. You can run the [ServerInventoryDataGathering.ps1](https://github.com/proxb/ServerInventoryReport/blob/master/ServerInventoryDataGathering.ps1) prior to implementing it as a scheduled task to enusre that there are no access issues both with the data gathering and the shipping of data to the SQL server. Again, you will need to update the script to point to the SQL server.

4. Once you are sure that everything looks good, you can create a scheduled task to run this script nightly or whenever you feel is appropriate.

5. Ensure that [ServerInventoryUI.ps1](https://github.com/proxb/ServerInventoryReport/blob/master/ServerInventoryUI.ps1) has been updated to point to your SQL server and also make sure that you have allowed the users that will be using the UI has read access to the database to allow the UI to populate the datatables correctly.
