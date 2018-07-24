function mapNetworkDrives()

username = 'leibd@wudosis.wustl.edu';

pswd = 'B2July04!%';

sysLine = ['net use j: \\ortho01.wudosis.wustl.edu\researchlabs2 /user:' username ' ' pswd '/persistent:Yes'];
system(sysLine);

sysLine = ['net use n: \\ortho01.wudosis.wustl.edu\silva''slab2 /user:' username ' ' pswd '/persistent:Yes'];
system(sysLine);