# VMwareBackupVeeam
Script I wrote that was able to grab a virtual machine move it from 1 cluster to another cluster to than make a backup with Veeam and move it back once finished.

We had a host that was not able to be ran with the Veeam backup, which was just one server. Cheeky solution I wrote was to move this Veeam to a cluster host with a license and run a manual Veeam backup job to create a backup of this virtual machine. Once the backup is finished the virtual machine would be moved back to its original cluster host.
