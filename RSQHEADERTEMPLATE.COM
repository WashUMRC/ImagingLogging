$ DEFINE/user sys$output dk0:[microct.data.SAMPLE.MEASUREMENT]RSQHeader1.txt
$ ctheader dk0:[microct.data.SAMPLE.MEASUREMENT]CNUMBER.RSQ;1
$ EXIT
