#!/bin/sh
cd -- "$(dirname "$BASH_SOURCE")"
python3 mats.py
echo done
sleep 1