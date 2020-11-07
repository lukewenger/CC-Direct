# CC-Direct
A non-visual GUI to manage, transform and filter huge amounts of pointcloud data with the help of CloudCompare

CloudCompare installation in default installation folder needed for the scripts to run(Path:"C:\Program Files\CloudCompare")
Basically the tool just sets together wished CC Commands and generates a command line prompt which will invoke if prompted.
The Tool loops through all files in Folder n. You will be prompted for every File found. This prevents from system crahes due to huge loops(for example wrong directory).

Combine Scans
This command works slightly different. Every scan in the directory will be taken and combined in one *.Bin File !!!!Filesize can be an issue here!!!!!!
If you are to combine Scans, only run this command, no queueing with other commands.
