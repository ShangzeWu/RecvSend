#! /bin/sh

find /var/www/html/RecvSend/uploadA/*  -name "*.xlsx*" -exec mv {} /var/www/html/backupFiles/Recv/uploadA \;

find /var/www/html/RecvSend/uploadB/*  -name "*.xlsx*" -exec mv {} /var/www/html/backupFiles/Recv/uploadB \;

find /var/www/html/RecvSend/uploadC/*  -name "*.xlsx*" -exec mv {} /var/www/html/backupFiles/Recv/uploadC \;

find /var/www/html/RecvSend/uploadD/*  -name "*.xlsx*" -exec mv {} /var/www/html/backupFiles/Recv/uploadD \;
