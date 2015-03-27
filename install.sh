#!/bin/sh

apt-get -y install python3-setuptools
python3 -m easy_install pip

python3 -m pip install eventlet
python3 -m pip install greenlet
python3 -m pip install contexttimer
