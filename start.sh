#!/bin/bash
cd /home/z/my-project
while true; do
    node .next/standalone/server.js 2>&1
    echo "Server exited, restarting in 2s..."
    sleep 2
done
