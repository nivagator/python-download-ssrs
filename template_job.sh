#!/bin/bash

echo "start"

# connect to OpenVPN
systemctl start openvpn@client.service
sleep 10
if ip tuntap | grep one_queue > /dev/null;
then
	echo "connected to vpn"
else
	echo "no tuntap"
fi

# execute py_dl_ssrs.main()
python3 /path/to/py_dl_ssrs.py
echo "py_dl_ssrs completed"

# disconnect from vpn
systemctl stop openvpn@client.service
echo "disconnected from vpn"

# move files
if mount | grep nas-mount-point > /dev/null;
then 
	echo "yay"
	rsync --remove-source-files -t /path/to/report_exports/* /path/to/mount/point/reports/
	rsync --remove-source-files -t /path/to/logs/* /path/to/mount/point/logs/
	echo "files moved"
else
	echo "NAS not detected"
fi 
echo "end"
