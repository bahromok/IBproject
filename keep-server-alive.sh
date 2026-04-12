#!/bin/bash
cd /home/z/my-project
while true; do
    node .next/standalone/server.js 2>>/tmp/next-errors.log
    echo "[$(date)] Server exited, restarting in 2s..." >> /tmp/next-errors.log
    sleep 2
done
